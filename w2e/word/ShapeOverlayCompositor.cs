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

            /* テキストを含む図形は、実際のテキストサイズに合わせて図形の大きさを補正する
             * （右端の文字切れや、図形の高さに対してテキストが小さすぎる場合の余白を防ぐため）
             */
            foreach( ShapeInfo shape in shapeInfoList )
            {
                FitShapeToContent( shape, scaleX, scaleY );
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
                    try
                    {
                        DrawShape( dc, shape, scaleX, scaleY, originOffsetXPx, originOffsetYPx, themeColors );
                    }
                    catch
                    {
                        /* 1つの図形の描画に失敗しても、他の図形・画像の合成自体は継続する */
                    }
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

            WordImageData result = ComposeFromShapeInfoList( a_mainDocumentPart, shapeInfoList );
            if( null != result )
            {
                result.relationshipId = "composed-shapes:" + a_paragraph.GetHashCode();
            }

            return result;
        }


        /// <summary>
        /// 解析済みの図形情報一覧から、1枚の画像に合成する（画像を伴わない、図形のみのケースで使用する共通処理）。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart（テーマ配色の解決に使用する）</param>
        /// <param name="a_shapeInfoList">合成対象の図形情報一覧（呼び出し側で座標調整済みのもの）</param>
        /// <returns>合成した画像情報。図形が無い場合はnull</returns>
        private static WordImageData ComposeFromShapeInfoList( MainDocumentPart a_mainDocumentPart, List<ShapeInfo> a_shapeInfoList )
        {
            if( null == a_shapeInfoList || 0 == a_shapeInfoList.Count )
            {
                return null;
            }

            /* 図形はベクター情報のみで解像度の基準となる画像が無いため、96DPI（1ピクセル=9525EMU）を基準に描画する。
             * これはWord/Office製品が既定で使用する画面表示解像度で、実寸に近いサイズで表示される。
             */
            double scaleX = 1.0 / EMU_PER_PIXEL;
            double scaleY = 1.0 / EMU_PER_PIXEL;

            /* テキストを含む図形は、実際のテキストサイズに合わせて図形の大きさを補正する
             * （右端の文字切れや、図形の高さに対してテキストが小さすぎる場合の余白を防ぐため）
             */
            foreach( ShapeInfo shape in a_shapeInfoList )
            {
                FitShapeToContent( shape, scaleX, scaleY );
            }

            double unionMinXEmu = double.MaxValue;
            double unionMinYEmu = double.MaxValue;
            double unionMaxXEmu = double.MinValue;
            double unionMaxYEmu = double.MinValue;

            foreach( ShapeInfo shape in a_shapeInfoList )
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
                foreach( ShapeInfo shape in a_shapeInfoList )
                {
                    try
                    {
                        DrawShape( dc, shape, scaleX, scaleY, originOffsetXPx, originOffsetYPx, themeColors );
                    }
                    catch
                    {
                        /* 1つの図形の描画に失敗しても、他の図形の合成自体は継続する */
                    }
                }
            }

            byte[] composedBytes = RenderToPng( visual, canvasWidth, canvasHeight );

            WordImageData result = new WordImageData();
            result.imageData = composedBytes;
            result.contentType = "image/png";
            result.relationshipId = "composed-shapes:" + Guid.NewGuid();
            result.widthEmu = (long)Math.Round( unionMaxXEmu - unionMinXEmu );
            result.heightEmu = (long)Math.Round( unionMaxYEmu - unionMinYEmu );
            result.altText = string.Empty;

            return result;
        }


        /// <summary>
        /// 段落をまたいで離れた位置に配置されている「図形のみ」の段落群を検出し、1枚の画像にまとめて合成する。
        /// Wordでは図形の位置は段落単位の相対座標（縦方向はrelativeFrom="paragraph"）で管理されているため、
        /// 正確な段落間の距離（実際のレイアウト上の高さ）は取得できない。そのため、間にある空段落の分だけ
        /// 既定の行高さ相当を積み上げる近似計算で縦位置を補正する（Wordの表示と完全には一致しない場合がある）。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_elements">Word本文の要素一覧</param>
        /// <param name="a_startIndex">走査を開始する段落のインデックス（図形のみの段落であること）</param>
        /// <param name="a_consumedCount">合成のために消費した（読み飛ばした）要素数。呼び出し側はこの数だけループを進める</param>
        /// <returns>合成した画像情報。対象外の場合はnull</returns>
        public static WordImageData TryComposeAcrossParagraphs( MainDocumentPart a_mainDocumentPart, IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> a_elements, int a_startIndex, out int a_consumedCount )
        {
            a_consumedCount = 1;

            /* Wordの標準的な1行分の高さの目安（EMU）。間にある空段落1つあたり、この高さ分だけ
             * 後続の図形の縦位置を押し下げる近似値として使用する
             */
            const double DEFAULT_LINE_HEIGHT_EMU = 190500;

            /* 図形同士の間に確保する余白（EMU） */
            const double GAP_EMU = 60000;

            try
            {
                if( !( a_elements[a_startIndex] is Word.Paragraph startParagraph ) )
                {
                    return null;
                }

                List<ShapeInfo> combinedShapes = new List<ShapeInfo>();
                double cumulativeYEmu = 0;
                int paragraphsWithShapes = 0;

                bool CollectParagraphShapes( Word.Paragraph a_para )
                {
                    List<Word.Drawing> drawings = a_para.Descendants<Word.Drawing>().ToList();
                    bool hasPicture = drawings.Any( d => null != WordImageHelper.GetBlip( d )?.Embed );

                    if( hasPicture )
                    {
                        /* 実際の画像（Blipを持つDrawing）があるパラグラフは対象外とする */
                        return false;
                    }

                    List<Word.Drawing> shapeDrawings = drawings.Where( IsShapeDrawing ).ToList();
                    if( 0 == shapeDrawings.Count )
                    {
                        return true;
                    }

                    List<ShapeInfo> shapes = ParseShapes( shapeDrawings );
                    if( 0 == shapes.Count )
                    {
                        return true;
                    }

                    double maxBottomEmu = 0;
                    foreach( ShapeInfo shape in shapes )
                    {
                        shape.YEmu += cumulativeYEmu;
                        maxBottomEmu = Math.Max( maxBottomEmu, shape.YEmu + shape.CyEmu );
                        combinedShapes.Add( shape );
                    }

                    cumulativeYEmu = maxBottomEmu + GAP_EMU;
                    paragraphsWithShapes++;
                    return true;
                }

                /* 起点となる段落自身の図形を取得する */
                if( !CollectParagraphShapes( startParagraph ) || 0 == combinedShapes.Count )
                {
                    return null;
                }

                /* 後続の段落を走査し、間に空段落を挟みつつ続く図形をまとめて取り込む */
                int scanIndex = a_startIndex + 1;
                while( scanIndex < a_elements.Count )
                {
                    if( !( a_elements[scanIndex] is Word.Paragraph scanParagraph ) )
                    {
                        /* 段落以外の要素（表など）が現れたらそこで走査を打ち切る */
                        break;
                    }

                    List<Word.Drawing> scanDrawings = scanParagraph.Descendants<Word.Drawing>().ToList();
                    bool scanHasPicture = scanDrawings.Any( d => null != WordImageHelper.GetBlip( d )?.Embed );
                    bool scanHasShape = scanDrawings.Any( IsShapeDrawing );

                    if( scanHasPicture )
                    {
                        /* 実際の画像が現れたらそこで走査を打ち切る（通常の画像処理に委ねる） */
                        break;
                    }

                    string scanText = WordHelper.GetVisibleText( scanParagraph )?.Trim();

                    if( !scanHasShape && !string.IsNullOrEmpty( scanText ) )
                    {
                        /* 実際の本文テキストが現れたらそこで走査を打ち切る */
                        break;
                    }

                    if( !scanHasShape )
                    {
                        /* 空段落は、既定の行高さ分だけ縦位置を押し下げる余白として扱う */
                        cumulativeYEmu += DEFAULT_LINE_HEIGHT_EMU;
                        scanIndex++;
                        continue;
                    }

                    CollectParagraphShapes( scanParagraph );
                    scanIndex++;
                }

                a_consumedCount = scanIndex - a_startIndex;

                /* 図形を含む段落が1つだけだった場合は、まとめる意味が無いため対象外とする
                 * （呼び出し側の通常の単一段落処理に任せる）
                 */
                if( 2 > paragraphsWithShapes )
                {
                    a_consumedCount = 1;
                    return null;
                }

                return ComposeFromShapeInfoList( a_mainDocumentPart, combinedShapes );
            }
            catch
            {
                a_consumedCount = 1;
                return null;
            }
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
        /// テキストを含む図形について、実際に描画されるテキストの幅・高さを測定し、
        /// 図形の実効サイズ（CxEmu/CyEmu）をそれに合わせて補正する。
        /// 幅は「足りなければ広げる」（右端の文字切れを防ぐ）、高さは「実測値に合わせる」
        /// （declaredされたサイズが実際のテキストより大きい場合の余分な空白を除去する）。
        /// テキストを持たない図形は何もしない。
        /// </summary>
        /// <param name="a_shape">対象の図形情報（呼び出し後、CxEmu/CyEmuが更新される）</param>
        /// <param name="a_scaleX">EMU→ピクセルの水平方向の倍率</param>
        /// <param name="a_scaleY">EMU→ピクセルの垂直方向の倍率</param>
        private static void FitShapeToContent( ShapeInfo a_shape, double a_scaleX, double a_scaleY )
        {
            if( null == a_shape.TextLines || 0 == a_shape.TextLines.Count )
            {
                return;
            }

            try
            {
                const double HORIZONTAL_PADDING_PX = 3;
                const double VERTICAL_PADDING_PX = 2;

                /* 末尾の空行（単なる段落終端マーカーで、Word上でも余分な高さとしては表示されないことが多い）は、
                 * 高さの計算対象から除外する。文中の空行（コード中の空白行など）は引き続き高さに含める。
                 */
                int lastNonEmptyIndex = -1;
                for( int i = 0; i < a_shape.TextLines.Count; i++ )
                {
                    if( 0 < a_shape.TextLines[i].Runs.Count )
                    {
                        lastNonEmptyIndex = i;
                    }
                }

                double maxWidthPx = 0;
                double totalHeightPx = VERTICAL_PADDING_PX;

                for( int i = 0; i <= lastNonEmptyIndex; i++ )
                {
                    double lineWidthPx = MeasureTextLine( a_shape.TextLines[i], a_scaleY, out double lineHeightPx );
                    maxWidthPx = Math.Max( maxWidthPx, lineWidthPx );
                    totalHeightPx += lineHeightPx;
                }

                totalHeightPx += VERTICAL_PADDING_PX;
                maxWidthPx += HORIZONTAL_PADDING_PX * 2;

                double neededCxEmu = maxWidthPx / a_scaleX;
                double neededCyEmu = totalHeightPx / a_scaleY;

                /* 幅は元の図形より狭くはしない（意図的な右余白を保つため）。高さは実測値に合わせて詰める */
                a_shape.CxEmu = Math.Max( a_shape.CxEmu, neededCxEmu );
                a_shape.CyEmu = neededCyEmu;
            }
            catch
            {
                /* 測定に失敗した場合は、元の図形サイズのまま処理を続行する */
            }
        }


        /// <summary>
        /// 図形内テキスト1行分の幅・高さ（ピクセル）を測定する（DrawShapeTextLinesと同じフォント設定ロジックを使用する）。
        /// </summary>
        /// <param name="a_line">測定対象の行</param>
        /// <param name="a_scaleY">EMU→ピクセルの垂直方向の倍率（フォントサイズの換算に使用する）</param>
        /// <param name="a_heightPx">その行の高さ（ピクセル）</param>
        /// <returns>その行の幅（ピクセル）</returns>
        private static double MeasureTextLine( ShapeTextLine a_line, double a_scaleY, out double a_heightPx )
        {
            const double EMU_PER_POINT = 12700.0;

            if( 0 == a_line.Runs.Count )
            {
                a_heightPx = 12.0 * EMU_PER_POINT * a_scaleY;
                return 0;
            }

            string combinedText = string.Concat( a_line.Runs.Select( r => r.Text ) );
            double firstFontSizePx = Math.Max( 4, a_line.Runs[0].FontSizePt * EMU_PER_POINT * a_scaleY );
            string defaultFontFamily = a_line.Runs[0].FontFamily ?? "Meiryo UI, Yu Gothic UI, Segoe UI";

            FormattedText measure = new FormattedText(
                combinedText,
                System.Globalization.CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface( new FontFamily( defaultFontFamily ), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal ),
                firstFontSizePx,
                Brushes.Black,
                96.0 );
            measure.MaxTextWidth = 1000000.0;

            int charIndex = 0;
            foreach( ShapeTextRun run in a_line.Runs )
            {
                int len = run.Text.Length;
                if( 0 == len ) { continue; }

                double runFontSizePx = Math.Max( 4, run.FontSizePt * EMU_PER_POINT * a_scaleY );
                measure.SetFontSize( runFontSizePx, charIndex, len );

                if( !string.IsNullOrEmpty( run.FontFamily ) )
                {
                    measure.SetFontFamily( new FontFamily( run.FontFamily ), charIndex, len );
                }
                if( run.Bold )
                {
                    measure.SetFontWeight( FontWeights.Bold, charIndex, len );
                }
                if( run.Italic )
                {
                    measure.SetFontStyle( FontStyles.Italic, charIndex, len );
                }

                charIndex += len;
            }

            a_heightPx = measure.Height;
            return measure.WidthIncludingTrailingWhitespace;
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
        /// 段落内の図形（wps:wsp）に設定されているテキストを一覧で取得する。
        /// Wordの文書には、図形に重なる位置に、図形内のテキストと同じ内容の
        /// 通常の段落（コピー&ペースト等の残骸で、図形の陰に隠れて表示されない段落）が
        /// 紛れ込んでいることがあるため、そのような重複段落を検出するために使用する。
        /// </summary>
        /// <param name="a_paragraph">対象段落</param>
        /// <returns>図形内に設定されているテキストの一覧（前後の空白は除去済み、空のものは含まない）</returns>
        public static List<string> GetShapeTexts( Word.Paragraph a_paragraph )
        {
            List<string> result = new List<string>();

            try
            {
                foreach( Word.Drawing drawing in a_paragraph.Descendants<Word.Drawing>() )
                {
                    if( !IsShapeDrawing( drawing ) )
                    {
                        continue;
                    }

                    ShapeInfo info = ParseShape( drawing );
                    string combinedText = CombineShapeText( info );
                    if( !string.IsNullOrEmpty( combinedText ) )
                    {
                        result.Add( combinedText );
                    }
                }
            }
            catch
            {
                /* 取得に失敗した場合は、重複判定を行わないよう空のリストを返す */
            }

            return result;
        }


        /// <summary>
        /// ShapeInfoのTextLines（行・run単位のテキスト）を、改行区切りの単純な文字列に結合する。
        /// 重複段落の検出や、外部からの一覧取得（GetShapeTexts）で使用する。
        /// </summary>
        private static string CombineShapeText( ShapeInfo a_info )
        {
            if( null == a_info?.TextLines )
            {
                return null;
            }

            IEnumerable<string> lineTexts = a_info.TextLines.Select( l => string.Concat( l.Runs.Select( r => r.Text ) ) );
            return string.Join( Environment.NewLine, lineTexts ).Trim();
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

            /// <summary>図形内のテキスト（行ごと・run（文字色や書式が変わる区間）ごとに保持する）。テキストが無い場合はnull</summary>
            public List<ShapeTextLine> TextLines;
        }


        /// <summary>
        /// 図形内テキストの1行分の情報
        /// </summary>
        private class ShapeTextLine
        {
            public List<ShapeTextRun> Runs = new List<ShapeTextRun>();

            /// <summary>行の配置。"left" / "center" / "right"。既定は"left"</summary>
            public string Alignment = "left";
        }


        /// <summary>
        /// 図形内テキストの1区間（Wordのrun）分の情報。文字色・フォント・サイズ・太字/斜体が変わるごとに分かれる
        /// </summary>
        private class ShapeTextRun
        {
            public string Text = "";
            public string ColorHex;
            public string FontFamily;
            public double FontSizePt = 10.5;
            public bool Bold;
            public bool Italic;
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
            XElement fillElement = spPr.Element( A_NS + "solidFill" );
            if( null != fillElement )
            {
                (info.FillColorHex, info.FillAlpha) = ParseColor( fillElement );
            }
            else if( null == spPr.Element( A_NS + "noFill" ) )
            {
                /* spPrに直接の塗り指定が無い場合、Word標準の「図形のスタイル」機能で選択された
                 * スタイル参照（wps:style/a:fillRef）から塗り色を解決する（明示的に「塗りつぶしなし」の場合を除く）
                 */
                XElement fillRef = wsp.Element( WPS_NS + "style" )?.Element( A_NS + "fillRef" );
                string fillRefIdx = fillRef?.Attribute( "idx" )?.Value;

                if( null != fillRef && "0" != fillRefIdx )
                {
                    (info.FillColorHex, info.FillAlpha) = ParseColor( fillRef );
                }
            }

            /* 線 */
            XElement ln = spPr.Element( A_NS + "ln" );
            if( null != ln )
            {
                info.LineWidthEmu = (double?)ln.Attribute( "w" ) ?? 0;
                (info.LineColorHex, info.LineAlpha) = ParseColor( ln.Element( A_NS + "solidFill" ) );
                info.HeadArrow = IsArrowMarker( ln.Element( A_NS + "headEnd" ) );
                info.TailArrow = IsArrowMarker( ln.Element( A_NS + "tailEnd" ) );
            }
            else
            {
                /* spPrに線の指定が無い場合も、同様にスタイル参照（wps:style/a:lnRef）からフォールバックする */
                XElement lnRef = wsp.Element( WPS_NS + "style" )?.Element( A_NS + "lnRef" );
                string lnRefIdx = lnRef?.Attribute( "idx" )?.Value;

                if( null != lnRef && "0" != lnRefIdx )
                {
                    (info.LineColorHex, info.LineAlpha) = ParseColor( lnRef );
                    /* スタイル参照のみで線幅が不明な場合は、細めの既定幅（0.75pt相当）を使用する */
                    info.LineWidthEmu = 9525;
                }
            }

            /* 図形内のテキスト（wps:txbx内の文字列）を取得する */
            ParseShapeText( wsp, info );

            return info;
        }


        /// <summary>
        /// 図形（wps:wsp）内のテキストボックス（wps:txbx）から、行ごと・run（文字色や書式の区間）ごとの
        /// テキスト・文字色・フォント・サイズ・配置を取得する。テキストが無い場合はShapeInfo.TextLinesをnullのままにする。
        /// </summary>
        private static void ParseShapeText( XElement a_wsp, ShapeInfo a_info )
        {
            XElement txbxContent = a_wsp?.Element( WPS_NS + "txbx" )?.Element( W_NS + "txbxContent" );
            if( null == txbxContent )
            {
                return;
            }

            List<ShapeTextLine> lines = new List<ShapeTextLine>();

            foreach( XElement paragraph in txbxContent.Elements( W_NS + "p" ) )
            {
                ShapeTextLine line = new ShapeTextLine();

                /* 段落の配置（左揃え・中央揃え・右揃え）。既定は左揃えとする */
                string jc = paragraph.Element( W_NS + "pPr" )?.Element( W_NS + "jc" )?.Attribute( W_NS + "val" )?.Value;
                if( "center" == jc || "right" == jc )
                {
                    line.Alignment = jc;
                }

                foreach( XElement run in paragraph.Elements( W_NS + "r" ) )
                {
                    string runText = string.Concat( run.Elements( W_NS + "t" ).Select( t => (string)t ) );
                    if( string.IsNullOrEmpty( runText ) )
                    {
                        continue;
                    }

                    XElement rPr = run.Element( W_NS + "rPr" );

                    ShapeTextRun textRun = new ShapeTextRun();
                    textRun.Text = runText;

                    /* 文字色 */
                    string colorVal = rPr?.Element( W_NS + "color" )?.Attribute( W_NS + "val" )?.Value;
                    if( !string.IsNullOrEmpty( colorVal ) && !"auto".Equals( colorVal, StringComparison.OrdinalIgnoreCase ) )
                    {
                        textRun.ColorHex = colorVal;
                    }

                    /* フォント名（半角文字用のasciiを使用する） */
                    string fontAscii = rPr?.Element( W_NS + "rFonts" )?.Attribute( W_NS + "ascii" )?.Value;
                    if( !string.IsNullOrEmpty( fontAscii ) )
                    {
                        textRun.FontFamily = fontAscii;
                    }

                    /* フォントサイズ（半ポイント単位） */
                    string sizeVal = rPr?.Element( W_NS + "sz" )?.Attribute( W_NS + "val" )?.Value;
                    if( !string.IsNullOrEmpty( sizeVal ) && double.TryParse( sizeVal, out double halfPoints ) )
                    {
                        textRun.FontSizePt = halfPoints / 2.0;
                    }

                    textRun.Bold = null != rPr?.Element( W_NS + "b" );
                    textRun.Italic = null != rPr?.Element( W_NS + "i" );

                    line.Runs.Add( textRun );
                }

                lines.Add( line );
            }

            /* すべての行が空（runが1つも無い）の場合はテキストなしとして扱う */
            if( lines.All( l => 0 == l.Runs.Count ) )
            {
                return;
            }

            a_info.TextLines = lines;
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
        /// a:solidFill要素から色(16進数)と不透明度を解析する。srgbClr / schemeClr / prstClr に対応する。
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
            XElement preset = a_solidFill.Element( A_NS + "prstClr" );

            XElement colorElement = srgb ?? scheme ?? preset;
            if( null == colorElement )
            {
                return (null, 1.0);
            }

            string val = colorElement.Attribute( "val" )?.Value;
            if( string.IsNullOrEmpty( val ) )
            {
                return (null, 1.0);
            }

            /* srgbClrの場合はそのまま16進数、schemeClrの場合は "scheme:" を、
             * prstClr（黒・白などの定義済み色名）の場合は "preset:" を付けてマークしておき、描画時に解決する
             */
            string colorKey;
            if( null != srgb )
            {
                colorKey = val;
            }
            else if( null != scheme )
            {
                colorKey = "scheme:" + val;
            }
            else
            {
                colorKey = "preset:" + val;
            }

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

            /* 図形内にテキストが設定されている場合は描画する（コネクタ／矢印には描画しない） */
            if( !isConnector && null != a_shape.TextLines && 0 < a_shape.TextLines.Count )
            {
                DrawShapeTextLines( a_dc, a_shape, bounds, a_scaleY, a_themeColors );
            }
        }


        /// <summary>
        /// 図形内のテキストを、行・run（文字色や書式の区間）ごとに、それぞれの配置・色・フォントで描画する。
        /// 上詰め・各行ごとの左揃え／中央揃え／右揃えに対応する（コードブロックのような複数行テキストを想定）。
        /// </summary>
        private static void DrawShapeTextLines( DrawingContext a_dc, ShapeInfo a_shape, Rect a_bounds, double a_scaleY, Dictionary<string, string> a_themeColors )
        {
            const double HORIZONTAL_PADDING_PX = 3;
            const double VERTICAL_PADDING_PX = 2;

            /* 1pt = 12700 EMU。フォントサイズ(pt)を、このキャンバスのEMU→ピクセル倍率でピクセルに変換する
             * （画像に重ねる場合は画像自身の解像度基準、図形のみの場合は96DPI基準の倍率になる）
             */
            const double EMU_PER_POINT = 12700.0;

            double y = a_bounds.Y + VERTICAL_PADDING_PX;

            foreach( ShapeTextLine line in a_shape.TextLines )
            {
                if( 0 == line.Runs.Count )
                {
                    /* 空行は既定サイズ相当だけ縦位置を進める */
                    y += 12.0 * EMU_PER_POINT * a_scaleY;
                    continue;
                }

                string combinedText = string.Concat( line.Runs.Select( r => r.Text ) );

                double firstFontSizePx = Math.Max( 4, line.Runs[0].FontSizePt * EMU_PER_POINT * a_scaleY );
                string defaultFontFamily = line.Runs[0].FontFamily ?? "Meiryo UI, Yu Gothic UI, Segoe UI";

                FormattedText formattedText = new FormattedText(
                    combinedText,
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface( new FontFamily( defaultFontFamily ), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal ),
                    firstFontSizePx,
                    Brushes.Black,
                    96.0 );

                /* 各Wordの段落を1行として扱い、途中で折り返さない（Word側の改行位置をそのまま尊重する）。
                 * double.PositiveInfinityを指定するとWPFの内部レイアウト計算で例外になることがあるため、
                 * 十分に大きい有限値を使用する
                 */
                formattedText.MaxTextWidth = 1000000.0;

                /* runごとに文字色・フォント・サイズ・太字/斜体を個別に適用する */
                int charIndex = 0;
                foreach( ShapeTextRun run in line.Runs )
                {
                    int len = run.Text.Length;
                    if( 0 == len ) { continue; }

                    Brush runBrush = ResolveBrush( run.ColorHex, 1.0, a_themeColors ) ?? Brushes.Black;
                    formattedText.SetForegroundBrush( runBrush, charIndex, len );

                    double runFontSizePx = Math.Max( 4, run.FontSizePt * EMU_PER_POINT * a_scaleY );
                    formattedText.SetFontSize( runFontSizePx, charIndex, len );

                    if( !string.IsNullOrEmpty( run.FontFamily ) )
                    {
                        formattedText.SetFontFamily( new FontFamily( run.FontFamily ), charIndex, len );
                    }
                    if( run.Bold )
                    {
                        formattedText.SetFontWeight( FontWeights.Bold, charIndex, len );
                    }
                    if( run.Italic )
                    {
                        formattedText.SetFontStyle( FontStyles.Italic, charIndex, len );
                    }

                    charIndex += len;
                }

                double x;
                if( "center" == line.Alignment )
                {
                    x = a_bounds.X + Math.Max( 0, ( a_bounds.Width - formattedText.WidthIncludingTrailingWhitespace ) / 2 );
                }
                else if( "right" == line.Alignment )
                {
                    x = a_bounds.X + Math.Max( HORIZONTAL_PADDING_PX, a_bounds.Width - formattedText.WidthIncludingTrailingWhitespace - HORIZONTAL_PADDING_PX );
                }
                else
                {
                    x = a_bounds.X + HORIZONTAL_PADDING_PX;
                }

                a_dc.DrawText( formattedText, new Point( x, y ) );

                y += formattedText.Height;
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
            else if( a_colorKey.StartsWith( "preset:", StringComparison.OrdinalIgnoreCase ) )
            {
                string presetName = a_colorKey.Substring( "preset:".Length );
                hex = ResolvePresetColorHex( presetName );
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
        /// DrawingMLのプリセット色名（a:prstClr、"black"や"white"等）を16進数カラーコードに変換する。
        /// よく使われる代表的な色名のみ対応し、未対応の色名はグレーにフォールバックする。
        /// </summary>
        private static string ResolvePresetColorHex( string a_presetName )
        {
            switch( ( a_presetName ?? "" ).ToLowerInvariant() )
            {
                case "black": return "000000";
                case "white": return "FFFFFF";
                case "red": return "FF0000";
                case "green": return "008000";
                case "blue": return "0000FF";
                case "yellow": return "FFFF00";
                case "orange": return "FFA500";
                case "purple": return "800080";
                case "gray":
                case "grey": return "808080";
                case "darkgray":
                case "darkgrey": return "A9A9A9";
                case "lightgray":
                case "lightgrey": return "D3D3D3";
                case "silver": return "C0C0C0";
                case "maroon": return "800000";
                case "navy": return "000080";
                case "teal": return "008080";
                case "olive": return "808000";
                case "lime": return "00FF00";
                case "aqua":
                case "cyan": return "00FFFF";
                case "magenta":
                case "fuchsia": return "FF00FF";
                case "brown": return "A52A2A";
                case "pink": return "FFC0CB";
                case "gold": return "FFD700";
                case "indigo": return "4B0082";
                case "violet": return "EE82EE";
                default: return "808080";
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
