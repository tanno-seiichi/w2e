using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
        private static readonly XNamespace W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

                if( 0 == shapeDrawings.Count )
                {
                    /* 重ねる図形が無ければ合成の意味が無いため対象外とする（通常の画像抽出に任せる） */
                    return null;
                }

                if( 1 == pictureDrawings.Count )
                {
                    /* 画像1枚 + 図形1つ以上 → 画像に図形を重ねて合成する */
                    return ComposeWithPicture( a_mainDocumentPart, a_paragraph, pictureDrawings[0], shapeDrawings );
                }

                if( 0 == pictureDrawings.Count )
                {
                    /* 画像が無く、図形のみ → 図形だけを1枚の画像として合成する */
                    return ComposeShapesOnly( a_mainDocumentPart, a_paragraph, shapeDrawings );
                }

                /* 画像が複数ある場合は、どの画像に図形が重なっているか特定できないため対象外とする */
                return null;
            }
            catch
            {
                /* 合成に失敗した場合は、通常の画像抽出処理にフォールバックする */
                return null;
            }
        }


        /// <summary>
        /// 画像1枚に、重なっている図形を合成する。
        /// </summary>
        private static WordImageData ComposeWithPicture( MainDocumentPart a_mainDocumentPart, Word.Paragraph a_paragraph, Word.Drawing a_pictureDrawing, List<Word.Drawing> a_shapeDrawings )
        {
            Blip pictureBlip = WordImageHelper.GetBlip( a_pictureDrawing );
            ImagePart imagePart = a_mainDocumentPart.GetPartById( pictureBlip.Embed.Value ) as ImagePart;

            if( null == imagePart )
            {
                return null;
            }

            byte[] baseImageBytes = WordImageHelper.GetImageData( imagePart );

            /* 画像自身の表示サイズ(EMU)を取得する（図形の位置をこの画像上の座標に変換するための基準にする） */
            WordImageHelper.GetImageSize( a_pictureDrawing, out long pictureWidthEmu, out long pictureHeightEmu );

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

            int imagePixelWidth = baseBitmap.PixelWidth;
            int imagePixelHeight = baseBitmap.PixelHeight;

            /* 図形の位置(EMU)を、画像のピクセル座標に変換するための倍率
             * （図形はWord上で画像と同じ段落基準の座標系に配置されているという前提で計算する）
             */
            double scaleX = imagePixelWidth / (double)pictureWidthEmu;
            double scaleY = imagePixelHeight / (double)pictureHeightEmu;

            /* 図形の情報を解析する */
            List<ShapeInfo> shapeInfoList = ParseShapes( a_shapeDrawings );

            if( 0 == shapeInfoList.Count )
            {
                /* 図形を1つも解析できなかった場合は合成の意味が無いため対象外とする */
                return null;
            }

            /* 図形の水平位置(relativeFrom="column")は段（ページの版面）を基準とした絶対座標だが、
             * 画像はインライン画像として段落内に配置されているため、段落に左インデントが設定されていると
             * 画像はその分だけ右にずれて表示される。図形の位置をこのインデント分だけ補正しないと、
             * 画像に対して図形が実際より右にずれて見えてしまう。
             */
            double indentEmu = GetLeftIndentEmu( a_paragraph );
            if( 0 != indentEmu )
            {
                foreach( ShapeInfo shape in shapeInfoList )
                {
                    shape.XEmu -= indentEmu;
                }
            }

            /* 図形は画像の外側にはみ出して配置されていることがある（Word上ではページの余白部分に
             * はみ出す形で違和感なく表示される）。はみ出した部分が切れないよう、画像と全ての図形を
             * 包含する外接矩形（EMU）を求め、その大きさに合わせてキャンバスを拡張する。
             */
            double unionMinXEmu = 0;
            double unionMinYEmu = 0;
            double unionMaxXEmu = pictureWidthEmu;
            double unionMaxYEmu = pictureHeightEmu;

            foreach( ShapeInfo shape in shapeInfoList )
            {
                unionMinXEmu = Math.Min( unionMinXEmu, shape.XEmu );
                unionMinYEmu = Math.Min( unionMinYEmu, shape.YEmu );
                unionMaxXEmu = Math.Max( unionMaxXEmu, shape.XEmu + shape.CxEmu );
                unionMaxYEmu = Math.Max( unionMaxYEmu, shape.YEmu + shape.CyEmu );
            }

            /* 画像原点(0,0)がキャンバス内のどこに来るかのオフセット（ピクセル） */
            double originOffsetXPx = -unionMinXEmu * scaleX;
            double originOffsetYPx = -unionMinYEmu * scaleY;

            int canvasWidth = (int)Math.Ceiling( ( unionMaxXEmu - unionMinXEmu ) * scaleX );
            int canvasHeight = (int)Math.Ceiling( ( unionMaxYEmu - unionMinYEmu ) * scaleY );

            /* テーマの配色（スキーム色の解決に使用） */
            Dictionary<string, string> themeColors = LoadThemeColors( a_mainDocumentPart );

            /* 描画する */
            DrawingVisual visual = new DrawingVisual();
            using( DrawingContext dc = visual.RenderOpen() )
            {
                dc.DrawImage( baseBitmap, new Rect( originOffsetXPx, originOffsetYPx, imagePixelWidth, imagePixelHeight ) );

                foreach( ShapeInfo shape in shapeInfoList )
                {
                    DrawShape( dc, shape, scaleX, scaleY, originOffsetXPx, originOffsetYPx, themeColors );
                }
            }

            byte[] composedBytes = RenderToPng( visual, canvasWidth, canvasHeight );

            WordImageData result = new WordImageData();
            result.imageData = composedBytes;
            result.contentType = "image/png";
            result.relationshipId = "composed:" + pictureBlip.Embed.Value;
            result.widthEmu = (long)Math.Round( unionMaxXEmu - unionMinXEmu );
            result.heightEmu = (long)Math.Round( unionMaxYEmu - unionMinYEmu );
            result.altText = string.Empty;

            return result;
        }


        /// <summary>
        /// 画像を伴わず、図形のみが配置されているケースで、図形だけを1枚の画像として合成する。
        /// </summary>
        private static WordImageData ComposeShapesOnly( MainDocumentPart a_mainDocumentPart, Word.Paragraph a_paragraph, List<Word.Drawing> a_shapeDrawings )
        {
            List<ShapeInfo> shapeInfoList = ParseShapes( a_shapeDrawings );

            if( 0 == shapeInfoList.Count )
            {
                return null;
            }

            /* 図形はベクター情報のみで解像度の基準となる画像が無いため、96DPI（1ピクセル=9525EMU）を基準に描画する。
             * これはWord/Office製品が既定で使用する画面表示解像度で、実寸に近いサイズで表示される。
             */
            double scaleX = 1.0 / EMU_PER_PIXEL;
            double scaleY = 1.0 / EMU_PER_PIXEL;

            /* 図形のみの場合、画像との位置合わせが不要なためインデント補正は行わず、
             * 図形同士の相対位置（段基準の絶対座標）をそのまま使用する。
             */
            double unionMinXEmu = double.MaxValue;
            double unionMinYEmu = double.MaxValue;
            double unionMaxXEmu = double.MinValue;
            double unionMaxYEmu = double.MinValue;

            foreach( ShapeInfo shape in shapeInfoList )
            {
                unionMinXEmu = Math.Min( unionMinXEmu, shape.XEmu );
                unionMinYEmu = Math.Min( unionMinYEmu, shape.YEmu );
                unionMaxXEmu = Math.Max( unionMaxXEmu, shape.XEmu + shape.CxEmu );
                unionMaxYEmu = Math.Max( unionMaxYEmu, shape.YEmu + shape.CyEmu );
            }

            double originOffsetXPx = -unionMinXEmu * scaleX;
            double originOffsetYPx = -unionMinYEmu * scaleY;

            int canvasWidth = Math.Max( 1, (int)Math.Ceiling( ( unionMaxXEmu - unionMinXEmu ) * scaleX ) );
            int canvasHeight = Math.Max( 1, (int)Math.Ceiling( ( unionMaxYEmu - unionMinYEmu ) * scaleY ) );

            Dictionary<string, string> themeColors = LoadThemeColors( a_mainDocumentPart );

            /* 背景は付けず透過のまま描画する（写真などの土台が無いため） */
            DrawingVisual visual = new DrawingVisual();
            using( DrawingContext dc = visual.RenderOpen() )
            {
                foreach( ShapeInfo shape in shapeInfoList )
                {
                    DrawShape( dc, shape, scaleX, scaleY, originOffsetXPx, originOffsetYPx, themeColors );
                }
            }

            byte[] composedBytes = RenderToPng( visual, canvasWidth, canvasHeight );

            WordImageData result = new WordImageData();
            result.imageData = composedBytes;
            result.contentType = "image/png";
            result.relationshipId = "composed-shapes:" + a_paragraph.GetHashCode();
            result.widthEmu = (long)Math.Round( unionMaxXEmu - unionMinXEmu );
            result.heightEmu = (long)Math.Round( unionMaxYEmu - unionMinYEmu );
            result.altText = string.Empty;

            return result;
        }


        /// <summary>
        /// DrawingのリストからShapeInfoを解析する（解析できなかった図形は結果に含めない）。
        /// </summary>
        private static List<ShapeInfo> ParseShapes( List<Word.Drawing> a_shapeDrawings )
        {
            List<ShapeInfo> shapeInfoList = new List<ShapeInfo>();

            foreach( Word.Drawing shapeDrawing in a_shapeDrawings )
            {
                ShapeInfo info = ParseShape( shapeDrawing );
                if( null != info )
                {
                    shapeInfoList.Add( info );
                }
            }

            return shapeInfoList;
        }


        /// <summary>
        /// DrawingVisualをPNGバイト列に変換する。
        /// </summary>
        private static byte[] RenderToPng( DrawingVisual a_visual, int a_width, int a_height )
        {
            RenderTargetBitmap rtb = new RenderTargetBitmap( a_width, a_height, 96, 96, PixelFormats.Pbgra32 );
            rtb.Render( a_visual );

            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add( BitmapFrame.Create( rtb ) );

            using( MemoryStream ms = new MemoryStream() )
            {
                encoder.Save( ms );
                return ms.ToArray();
            }
        }


        /// <summary>
        /// 段落に設定されている左インデントをEMU単位で取得する。
        /// 図形の位置（段基準の絶対座標）と、段落内のインライン画像の表示位置とのズレを補正するために使用する。
        /// </summary>
        /// <param name="a_paragraph">対象段落</param>
        /// <returns>左インデント（EMU）。設定が無い場合は0</returns>
        private static double GetLeftIndentEmu( Word.Paragraph a_paragraph )
        {
            /* 1 twip (1/20 pt) = 635 EMU */
            const double EMU_PER_TWIP = 635.0;

            Word.Indentation indentation = a_paragraph.ParagraphProperties?.Indentation;
            if( null == indentation )
            {
                return 0;
            }

            /* w:left（またはRTL対応のw:start）を優先する。数値変換できない場合は0とする */
            string twipsText = indentation.Left?.Value ?? indentation.Start?.Value;
            if( string.IsNullOrEmpty( twipsText ) )
            {
                return 0;
            }

            if( !double.TryParse( twipsText, out double twips ) )
            {
                return 0;
            }

            return twips * EMU_PER_TWIP;
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
            public string Text;
            public string TextColorHex;
            public double TextFontSizePt = 10.5;
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

            /* 図形内のテキスト（wps:txbx内の文字列）を取得する */
            ParseShapeText( wsp, info );

            return info;
        }


        /// <summary>
        /// 図形（wps:wsp）内のテキストボックス（wps:txbx）から、表示テキスト・文字色・フォントサイズを取得する。
        /// テキストが無い場合はShapeInfo.Textをnullのままにする。
        /// </summary>
        private static void ParseShapeText( XElement a_wsp, ShapeInfo a_info )
        {
            XElement txbxContent = a_wsp?.Element( WPS_NS + "txbx" )?.Element( W_NS + "txbxContent" );
            if( null == txbxContent )
            {
                return;
            }

            /* 段落ごとの文字列を改行で連結する */
            List<string> paragraphTexts = new List<string>();
            XElement firstRunProperties = null;

            foreach( XElement paragraph in txbxContent.Elements( W_NS + "p" ) )
            {
                StringBuilder paragraphText = new StringBuilder();

                foreach( XElement run in paragraph.Elements( W_NS + "r" ) )
                {
                    if( null == firstRunProperties )
                    {
                        firstRunProperties = run.Element( W_NS + "rPr" );
                    }

                    foreach( XElement textElement in run.Elements( W_NS + "t" ) )
                    {
                        paragraphText.Append( (string)textElement );
                    }
                }

                paragraphTexts.Add( paragraphText.ToString() );
            }

            string combinedText = string.Join( Environment.NewLine, paragraphTexts ).Trim();
            if( string.IsNullOrEmpty( combinedText ) )
            {
                return;
            }

            a_info.Text = combinedText;

            /* 文字色（先頭の実行から取得。無ければ既定色を描画時に使用する） */
            XElement colorElement = firstRunProperties?.Element( W_NS + "color" );
            string colorVal = colorElement?.Attribute( W_NS + "val" )?.Value;
            if( !string.IsNullOrEmpty( colorVal ) && !"auto".Equals( colorVal, StringComparison.OrdinalIgnoreCase ) )
            {
                a_info.TextColorHex = colorVal;
            }

            /* フォントサイズ（半ポイント単位） */
            XElement sizeElement = firstRunProperties?.Element( W_NS + "sz" );
            string sizeVal = sizeElement?.Attribute( W_NS + "val" )?.Value;
            if( !string.IsNullOrEmpty( sizeVal ) && double.TryParse( sizeVal, out double halfPoints ) )
            {
                a_info.TextFontSizePt = halfPoints / 2.0;
            }
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
        private static void DrawShape( DrawingContext a_dc, ShapeInfo a_shape, double a_scaleX, double a_scaleY, double a_offsetXPx, double a_offsetYPx, Dictionary<string, string> a_themeColors )
        {
            double x = a_shape.XEmu * a_scaleX + a_offsetXPx;
            double y = a_shape.YEmu * a_scaleY + a_offsetYPx;
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

            /* 図形内にテキストが設定されている場合は、図形の中央に描画する（コネクタ／矢印には描画しない） */
            if( !isConnector && !string.IsNullOrEmpty( a_shape.Text ) )
            {
                DrawShapeText( a_dc, a_shape, bounds, a_themeColors );
            }
        }


        /// <summary>
        /// 図形内のテキストを、図形の中央に収まるように描画する。
        /// </summary>
        private static void DrawShapeText( DrawingContext a_dc, ShapeInfo a_shape, Rect a_bounds, Dictionary<string, string> a_themeColors )
        {
            Brush textBrush = ResolveBrush( a_shape.TextColorHex, 1.0, a_themeColors ) ?? Brushes.Black;

            /* フォントサイズ(pt)をピクセルに変換する（1pt = 96/72 px） */
            double fontSizePx = a_shape.TextFontSizePt * ( 96.0 / 72.0 );
            if( 6 > fontSizePx )
            {
                fontSizePx = 6;
            }

            FormattedText formattedText = new FormattedText(
                a_shape.Text,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface( new FontFamily( "Meiryo UI, Yu Gothic UI, Segoe UI" ), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal ),
                fontSizePx,
                textBrush,
                96.0 );

            formattedText.TextAlignment = TextAlignment.Center;

            /* 図形の高さを超える場合は、収まるようにフォントサイズを縮小する */
            if( formattedText.Height > a_bounds.Height && 0 < a_bounds.Height )
            {
                double shrinkRatio = a_bounds.Height / formattedText.Height;
                double adjustedSizePx = Math.Max( 6, fontSizePx * shrinkRatio );

                formattedText = new FormattedText(
                    a_shape.Text,
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface( new FontFamily( "Meiryo UI, Yu Gothic UI, Segoe UI" ), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal ),
                    adjustedSizePx,
                    textBrush,
                    96.0 );
                formattedText.TextAlignment = TextAlignment.Center;
            }

            formattedText.MaxTextWidth = Math.Max( 1, a_bounds.Width );

            Point origin = new Point( a_bounds.X, a_bounds.Y + Math.Max( 0, ( a_bounds.Height - formattedText.Height ) / 2 ) );
            a_dc.DrawText( formattedText, origin );
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
