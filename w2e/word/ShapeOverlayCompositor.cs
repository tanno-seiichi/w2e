using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using Point = System.Windows.Point;

namespace w2e.word
{
    /// <summary>
    /// Word文書内で、画像の上に矢印や強調枠などの図形が重ねて配置されているケースを検出し、
    /// それらを1枚の画像として合成するクラス。
    /// Markdown / Excelは実際の画像ファイルしか埋め込めないため、
    /// 図形（ベクター描画オブジェクト）は画像に焼き込んでから出力する必要がある。
    /// </summary>
    public static class ShapeOverlayCompositor
    {
        private static readonly XNamespace WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private static readonly XNamespace A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        /// <summary>
        /// EMU（English Metric Units）から96DPI換算のピクセル数への変換係数
        /// </summary>
        private const double EMU_PER_PIXEL = 9525.0;


        /// <summary>
        /// 段落内に「1枚の画像」と「それに重なる図形」が存在する場合、1枚に合成した画像を生成する。
        /// 対象外の場合（画像が無い、図形が無い、画像が複数あるなど）はnullを返す。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_paragraph">対象段落</param>
        /// <returns>合成した画像情報。対象外の場合はnull</returns>
        public static WordImageData TryCompose( MainDocumentPart a_mainDocumentPart, Word.Paragraph a_paragraph )
        {
            try
            {
                List<Word.Drawing> drawings = a_paragraph.Descendants<Word.Drawing>().ToList();

                /* 画像（Blipを持つDrawing）と、図形のみのDrawingに分類する */
                List<Word.Drawing> pictureDrawings = new List<Word.Drawing>();
                List<Word.Drawing> shapeDrawings = new List<Word.Drawing>();

                foreach( Word.Drawing drawing in drawings )
                {
                    Blip blip = WordImageHelper.GetBlip( drawing );

                    if( null != blip && null != blip.Embed )
                    {
                        pictureDrawings.Add( drawing );
                    }
                    else if( IsShapeDrawing( drawing ) )
                    {
                        shapeDrawings.Add( drawing );
                    }
                }

                /* 画像が1枚、かつ重なる図形が1つ以上あるケースのみ合成対象とする */
                if( 1 != pictureDrawings.Count || 0 == shapeDrawings.Count )
                {
                    return null;
                }

                Word.Drawing pictureDrawing = pictureDrawings[0];
                Blip pictureBlip = WordImageHelper.GetBlip( pictureDrawing );
                ImagePart imagePart = a_mainDocumentPart.GetPartById( pictureBlip.Embed.Value ) as ImagePart;

                if( null == imagePart )
                {
                    return null;
                }

                byte[] baseImageBytes = WordImageHelper.GetImageData( imagePart );

                /* 画像自身の表示サイズ(EMU)を取得する（図形の位置をこの画像上の座標に変換するための基準にする） */
                WordImageHelper.GetImageSize( pictureDrawing, out long pictureWidthEmu, out long pictureHeightEmu );

                if( 0 >= pictureWidthEmu || 0 >= pictureHeightEmu )
                {
                    return null;
                }

                /* ベースとなる画像を読み込む */
                BitmapImage baseBitmap = new BitmapImage();
                using( MemoryStream ms = new MemoryStream( baseImageBytes ) )
                {
                    baseBitmap.BeginInit();
                    baseBitmap.CacheOption = BitmapCacheOption.OnLoad;
                    baseBitmap.StreamSource = ms;
                    baseBitmap.EndInit();
                }
                baseBitmap.Freeze();

                int canvasWidth = baseBitmap.PixelWidth;
                int canvasHeight = baseBitmap.PixelHeight;

                /* 図形の位置(EMU)を、この画像のピクセル座標に変換するための倍率
                 * （図形はWord上で画像と同じ段落基準の座標系に配置されているという前提で計算する）
                 */
                double scaleX = canvasWidth / (double)pictureWidthEmu;
                double scaleY = canvasHeight / (double)pictureHeightEmu;

                /* 図形の情報を解析する */
                List<ShapeInfo> shapeInfoList = new List<ShapeInfo>();
                foreach( Word.Drawing shapeDrawing in shapeDrawings )
                {
                    ShapeInfo info = ParseShape( shapeDrawing );
                    if( null != info )
                    {
                        shapeInfoList.Add( info );
                    }
                }

                if( 0 == shapeInfoList.Count )
                {
                    /* 図形を1つも解析できなかった場合は合成の意味が無いため対象外とする */
                    return null;
                }

                /* テーマの配色（スキーム色の解決に使用） */
                Dictionary<string, string> themeColors = LoadThemeColors( a_mainDocumentPart );

                /* 描画する */
                DrawingVisual visual = new DrawingVisual();
                using( DrawingContext dc = visual.RenderOpen() )
                {
                    dc.DrawImage( baseBitmap, new Rect( 0, 0, canvasWidth, canvasHeight ) );

                    foreach( ShapeInfo shape in shapeInfoList )
                    {
                        DrawShape( dc, shape, scaleX, scaleY, themeColors );
                    }
                }

                RenderTargetBitmap rtb = new RenderTargetBitmap( canvasWidth, canvasHeight, 96, 96, PixelFormats.Pbgra32 );
                rtb.Render( visual );

                PngBitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add( BitmapFrame.Create( rtb ) );

                byte[] composedBytes;
                using( MemoryStream ms = new MemoryStream() )
                {
                    encoder.Save( ms );
                    composedBytes = ms.ToArray();
                }

                WordImageData result = new WordImageData();
                result.imageData = composedBytes;
                result.contentType = "image/png";
                result.relationshipId = "composed:" + pictureBlip.Embed.Value;
                result.widthEmu = pictureWidthEmu;
                result.heightEmu = pictureHeightEmu;
                result.altText = string.Empty;

                return result;
            }
            catch
            {
                /* 合成に失敗した場合は、通常の画像抽出処理にフォールバックする */
                return null;
            }
        }


        /// <summary>
        /// Drawingが、画像を持たない「図形のみ」のDrawing（wps:wsp）かどうかを判定する。
        /// </summary>
        private static bool IsShapeDrawing( Word.Drawing a_drawing )
        {
            return a_drawing.Descendants().Any( e => "wsp" == e.LocalName );
        }


        /// <summary>
        /// 図形1つ分の情報
        /// </summary>
        private class ShapeInfo
        {
            public double XEmu;
            public double YEmu;
            public double CxEmu;
            public double CyEmu;
            public string Preset;
            public bool FlipH;
            public bool FlipV;
            public string FillColorHex;
            public double FillAlpha = 1.0;
            public string LineColorHex;
            public double LineAlpha = 1.0;
            public double LineWidthEmu;
            public bool HeadArrow;
            public bool TailArrow;
        }


        /// <summary>
        /// 図形のDrawingから位置・サイズ・塗り・線・矢印の情報を解析する。
        /// 位置情報（絶対オフセット）が取得できない図形（中央揃え等の相対配置）は非対応のためnullを返す。
        /// </summary>
        private static ShapeInfo ParseShape( Word.Drawing a_drawing )
        {
            XElement root = XElement.Parse( a_drawing.OuterXml );

            /* wp:anchor（フロート配置）のみ対応する。wp:inline（行内配置）の図形は重ね合わせ対象として扱わない */
            XElement anchor = root.Element( WP_NS + "anchor" );
            if( null == anchor )
            {
                return null;
            }

            XElement posOffsetH = anchor.Element( WP_NS + "positionH" )?.Element( WP_NS + "posOffset" );
            XElement posOffsetV = anchor.Element( WP_NS + "positionV" )?.Element( WP_NS + "posOffset" );
            XElement extent = anchor.Element( WP_NS + "extent" );

            if( null == posOffsetH || null == posOffsetV || null == extent )
            {
                /* 絶対位置が指定されていない（中央揃え等）図形は座標を特定できないため対象外とする */
                return null;
            }

            XElement wsp = anchor.Descendants( WPS_NS + "wsp" ).FirstOrDefault();
            XElement spPr = wsp?.Element( WPS_NS + "spPr" );

            if( null == spPr )
            {
                return null;
            }

            ShapeInfo info = new ShapeInfo();
            info.XEmu = (double?)posOffsetH ?? 0;
            info.YEmu = (double?)posOffsetV ?? 0;
            info.CxEmu = (double?)extent.Attribute( "cx" ) ?? 0;
            info.CyEmu = (double?)extent.Attribute( "cy" ) ?? 0;

            XElement xfrm = spPr.Element( A_NS + "xfrm" );
            info.FlipH = "1" == xfrm?.Attribute( "flipH" )?.Value;
            info.FlipV = "1" == xfrm?.Attribute( "flipV" )?.Value;

            info.Preset = spPr.Element( A_NS + "prstGeom" )?.Attribute( "prst" )?.Value ?? "rect";

            /* 塗りつぶし */
            (info.FillColorHex, info.FillAlpha) = ParseColor( spPr.Element( A_NS + "solidFill" ) );

            /* 線 */
            XElement ln = spPr.Element( A_NS + "ln" );
            if( null != ln )
            {
                info.LineWidthEmu = (double?)ln.Attribute( "w" ) ?? 0;
                (info.LineColorHex, info.LineAlpha) = ParseColor( ln.Element( A_NS + "solidFill" ) );
                info.HeadArrow = IsArrowMarker( ln.Element( A_NS + "headEnd" ) );
                info.TailArrow = IsArrowMarker( ln.Element( A_NS + "tailEnd" ) );
            }

            return info;
        }


        /// <summary>
        /// a:headEnd / a:tailEnd が矢印マーカーを表しているかどうかを判定する
        /// </summary>
        private static bool IsArrowMarker( XElement a_markerElement )
        {
            string type = a_markerElement?.Attribute( "type" )?.Value;
            return !string.IsNullOrEmpty( type ) && "none" != type;
        }


        /// <summary>
        /// a:solidFill要素から色(16進数)と不透明度を解析する。srgbClr / schemeClr の両方に対応する。
        /// スキーム色（テーマ色）の解決自体はここでは行わず、色名の文字列（"accent2"等）をそのまま返す。
        /// 実際のRGBへの解決は描画時にテーマ配色を使って行う。
        /// </summary>
        private static (string colorHexOrSchemeName, double alpha) ParseColor( XElement a_solidFill )
        {
            if( null == a_solidFill )
            {
                return (null, 1.0);
            }

            XElement srgb = a_solidFill.Element( A_NS + "srgbClr" );
            XElement scheme = a_solidFill.Element( A_NS + "schemeClr" );

            XElement colorElement = srgb ?? scheme;
            if( null == colorElement )
            {
                return (null, 1.0);
            }

            string val = colorElement.Attribute( "val" )?.Value;
            if( string.IsNullOrEmpty( val ) )
            {
                return (null, 1.0);
            }

            /* srgbClrの場合はそのまま16進数、schemeClrの場合は "scheme:" を付けてマークしておき、描画時に解決する */
            string colorKey = ( null != srgb ) ? val : "scheme:" + val;

            double alpha = 1.0;
            XElement alphaElement = colorElement.Element( A_NS + "alpha" );
            if( null != alphaElement )
            {
                double alphaVal = (double?)alphaElement.Attribute( "val" ) ?? 100000.0;
                alpha = alphaVal / 100000.0;
            }

            return (colorKey, alpha);
        }


        /// <summary>
        /// 図形1つを描画する。
        /// </summary>
        private static void DrawShape( DrawingContext a_dc, ShapeInfo a_shape, double a_scaleX, double a_scaleY, Dictionary<string, string> a_themeColors )
        {
            double x = a_shape.XEmu * a_scaleX;
            double y = a_shape.YEmu * a_scaleY;
            double w = a_shape.CxEmu * a_scaleX;
            double h = a_shape.CyEmu * a_scaleY;
            Rect bounds = new Rect( x, y, Math.Max( 0, w ), Math.Max( 0, h ) );

            Brush fillBrush = ResolveBrush( a_shape.FillColorHex, a_shape.FillAlpha, a_themeColors );

            double lineWidthPx = a_shape.LineWidthEmu / EMU_PER_PIXEL;
            Brush lineBrush = ResolveBrush( a_shape.LineColorHex, a_shape.LineAlpha, a_themeColors );
            Pen pen = ( null != lineBrush && 0 < lineWidthPx ) ? new Pen( lineBrush, lineWidthPx ) : null;

            bool isConnector = null != a_shape.Preset &&
                                ( a_shape.Preset.IndexOf( "Connector", StringComparison.OrdinalIgnoreCase ) >= 0 ||
                                  "line" == a_shape.Preset );

            if( isConnector )
            {
                Point p1 = new Point( a_shape.FlipH ? bounds.Right : bounds.Left, a_shape.FlipV ? bounds.Bottom : bounds.Top );
                Point p2 = new Point( a_shape.FlipH ? bounds.Left : bounds.Right, a_shape.FlipV ? bounds.Top : bounds.Bottom );

                Pen linePen = pen ?? new Pen( Brushes.Black, 1.0 );
                a_dc.DrawLine( linePen, p1, p2 );

                Brush arrowBrush = lineBrush ?? Brushes.Black;
                if( a_shape.HeadArrow )
                {
                    DrawArrowhead( a_dc, p2, p1, arrowBrush, Math.Max( 4, lineWidthPx * 3 ) );
                }
                if( a_shape.TailArrow )
                {
                    DrawArrowhead( a_dc, p1, p2, arrowBrush, Math.Max( 4, lineWidthPx * 3 ) );
                }
            }
            else if( "ellipse" == a_shape.Preset )
            {
                a_dc.DrawEllipse( fillBrush, pen, new Point( bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2 ), bounds.Width / 2, bounds.Height / 2 );
            }
            else
            {
                /* rect, roundRect およびその他未対応の図形は矩形として描画する */
                a_dc.DrawRectangle( fillBrush, pen, bounds );
            }
        }


        /// <summary>
        /// 線分の終点(a_to)に矢印の三角形を描画する。a_fromからa_toへ向かう向きに合わせる。
        /// </summary>
        private static void DrawArrowhead( DrawingContext a_dc, Point a_from, Point a_to, Brush a_brush, double a_size )
        {
            double dx = a_to.X - a_from.X;
            double dy = a_to.Y - a_from.Y;
            double len = Math.Sqrt( dx * dx + dy * dy );
            if( 0 >= len )
            {
                return;
            }

            double ux = dx / len;
            double uy = dy / len;
            double px = -uy;
            double py = ux;

            Point baseCenter = new Point( a_to.X - ux * a_size, a_to.Y - uy * a_size );
            Point wing1 = new Point( baseCenter.X + px * a_size * 0.5, baseCenter.Y + py * a_size * 0.5 );
            Point wing2 = new Point( baseCenter.X - px * a_size * 0.5, baseCenter.Y - py * a_size * 0.5 );

            StreamGeometry geometry = new StreamGeometry();
            using( StreamGeometryContext ctx = geometry.Open() )
            {
                ctx.BeginFigure( a_to, true, true );
                ctx.LineTo( wing1, true, true );
                ctx.LineTo( wing2, true, true );
            }
            geometry.Freeze();

            a_dc.DrawGeometry( a_brush, null, geometry );
        }


        /// <summary>
        /// 色情報（16進数、または "scheme:" で始まるスキーム色名）からBrushを解決する。
        /// </summary>
        private static Brush ResolveBrush( string a_colorKey, double a_alpha, Dictionary<string, string> a_themeColors )
        {
            if( string.IsNullOrEmpty( a_colorKey ) )
            {
                return null;
            }

            string hex = a_colorKey;
            if( a_colorKey.StartsWith( "scheme:", StringComparison.OrdinalIgnoreCase ) )
            {
                string schemeName = a_colorKey.Substring( "scheme:".Length );
                if( !a_themeColors.TryGetValue( schemeName, out hex ) )
                {
                    /* テーマ色が解決できない場合は既定のグレーにフォールバックする */
                    hex = "808080";
                }
            }

            try
            {
                Color color = (Color)ColorConverter.ConvertFromString( "#" + hex );
                color.A = (byte)Math.Round( Math.Max( 0.0, Math.Min( 1.0, a_alpha ) ) * 255 );

                SolidColorBrush brush = new SolidColorBrush( color );
                brush.Freeze();
                return brush;
            }
            catch
            {
                return null;
            }
        }


        /// <summary>
        /// テーマ（配色）の定義を読み込み、スキーム色名から16進数カラーコードへのマップを作成する。
        /// テーマが取得できない場合は空のマップを返す（呼び出し側で既定色にフォールバックする）。
        /// </summary>
        private static Dictionary<string, string> LoadThemeColors( MainDocumentPart a_mainDocumentPart )
        {
            Dictionary<string, string> map = new Dictionary<string, string>( StringComparer.OrdinalIgnoreCase );

            try
            {
                ThemePart themePart = a_mainDocumentPart.ThemePart;
                if( null == themePart )
                {
                    return map;
                }

                XElement themeXml;
                using( Stream stream = themePart.GetStream() )
                {
                    themeXml = XElement.Load( stream );
                }

                XElement clrScheme = themeXml.Descendants( A_NS + "clrScheme" ).FirstOrDefault();
                if( null == clrScheme )
                {
                    return map;
                }

                string[] names = { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
                foreach( string name in names )
                {
                    XElement el = clrScheme.Element( A_NS + name );
                    string val = el?.Element( A_NS + "srgbClr" )?.Attribute( "val" )?.Value
                                 ?? el?.Element( A_NS + "sysClr" )?.Attribute( "lastClr" )?.Value;

                    if( !string.IsNullOrEmpty( val ) )
                    {
                        map[name] = val;
                    }
                }
            }
            catch
            {
                /* テーマの取得に失敗しても、既定色へのフォールバックで処理を継続する */
            }

            return map;
        }
    }
}
