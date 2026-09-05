using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using w2e.delegates;
using w2e.file;
using w2e.markdown;
using w2e.word;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace w2e.converter
{
    /// <summary>
    /// Wordファイルの内容をMarkDownファイルに書き出すクラス
    /// </summary>
    public class W2MdConverter : IConverter
    {
        /// <summary>
        /// 開始時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_MIN_VALUE = 1;

        /// <summary>
        /// Word読込完了時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_WORD_RANGE = 40;

        /// <summary>
        /// MarkDown書出完了時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_MARKDOWN_RANGE = 60;

        /// <summary>
        /// 先頭ファイルのファイル名
        /// </summary>
        private const string TOP_FILE_NAME = "0 トップ.md";

        /// <summary>
        /// 画像ファイル名を生成するオブジェクト
        /// </summary>
        private readonly ImageFileNameGenerator m_imageFileNameGenerator = new ImageFileNameGenerator();

        /// <summary>
        /// 進捗情報が更新された時の処理
        /// </summary>
        public Delegates.UpdateProgressDelegate onProgressUpdate { get; set; }

        /// <summary>
        /// ログが出力された時の処理
        /// </summary>
        public Delegates.UpdateLogDelegate onLogUpdate { get; set; }


        /// <summary>
        /// Word → MarkDown の変換処理を実行する
        /// </summary>
        /// <param name="a_wordPath">Wordファイルのパス</param>
        /// <param name="a_outputDir">MarkDownファイルの出力先ディレクトリ</param>
        /// <param name="a_outputImage_flg">画像を使用するか否か</param>
        /// <param name="a_outputListNumber_flg">箇条書きに番号を使用するか否か。falseの場合は固定で「-」を使用する</param>
        /// <param name="a_token">処理中断通知</param>
        public void Convert( string a_wordPath, string a_outputDir, bool a_outputImage_flg, bool a_outputListNumber_flg, CancellationToken a_token )
        {
            onProgressUpdate?.Invoke( PROGRESS_MIN_VALUE );
            string tempPath = FileCopy.CreateTempCopy(a_wordPath);

            try
            {
                /* Wordファイルを読込 */
                using( WordprocessingDocument doc = WordprocessingDocument.Open( tempPath, false ) )
                {
                    Word.Body body = doc.MainDocumentPart.Document.Body;
                    StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;

                    Dictionary<int, NumberingDefinition> numberingMap = WordHelper.LoadNumbring(doc);
                    NumberingEngine engine = new NumberingEngine();

                    /* MarkDownWriterのインスタンスを生成 */
                    var md = new MarkDownWriter();

                    /* 現在の章番号を初期化(画像ファイルのファイル名に使用する) */
                    string currentNum = string.Empty;

                    /* 直前に新規ファイルとして採用した章タイトルを保持する
                     * （見出しテキストが空で、次の非空段落をタイトルとして補完する際に、
                     *   その補完先が直前の章タイトルと同一だった場合は、
                     *   実体のない空の見出し段落（コピー&ペースト等で紛れ込んだもの）を
                     *   誤って新しい章として分割してしまわないようにするために使用する）
                     */
                    string lastChapterTitle = null;

                    /* 現在の章内で出現した見出しの階層スタック（章見出し自身を深さ1として、以降に出現した見出しの
                     * 相対的な深さを「直前に出現した見出しとの関係」から決めるために使用する）
                     * 各要素は (Word上のレベル, 割り当てた深さ) のペア
                     */
                    var headingStack = new List<(int Level, int Depth)>();

                    /* 直前に処理した図形（矢印・強調枠等）の中に設定されていたテキストを保持する。
                     * Word上でその図形の陰に隠れて表示されない「重複した通常の段落」を検出するために使用する
                     * （詳細は使用箇所のコメントを参照）
                     */
                    var recentShapeTexts = new HashSet<string>();

                    /* 現在のファイル情報を初期化 */
                    string fileName = TOP_FILE_NAME;
                    string filePath = Path.Combine( a_outputDir, fileName );
                    md.NewFile( filePath );

                    /* ログにファイル名を表示 */
                    onLogUpdate( fileName );

                    int total = body.Elements().Count();
                    int current = 0;

                    /* プログレスバーをWordファイル読込終了まで進める */
                    onProgressUpdate?.Invoke( PROGRESS_WORD_RANGE );

                    /* Wordファイルの要素一覧（見出しのタイトル補完のため、次要素を先読みできるようリスト化する） */
                    List<OpenXmlElement> elements = body.Elements().ToList();

                    /* Wordファイルの要素ごとに処理 */
                    for( int elementIndex = 0; elementIndex < elements.Count; elementIndex++ )
                    {
                        OpenXmlElement element = elements[elementIndex];

                        /* 処理中断が要求されていたらループを抜ける */
                        if( a_token.IsCancellationRequested ) break;

                        /* プログレスバーを更新 */
                        current++;
                        int progress = PROGRESS_WORD_RANGE + (int)(current * PROGRESS_MARKDOWN_RANGE / total);
                        onProgressUpdate?.Invoke( progress );

                        /* Wordファイル「段落」の処理 */
                        Word.Paragraph para = element as Word.Paragraph;
                        if( null != para )
                        {
                            var info = WordHelper.GetNumberingInfo(para, stylePart);
                            int? numId = info.numId;
                            int? level = info.level;
                            WordHelper.NumberingTypeEn numberingType = info.numberingType;

                            string text = WordHelper.GetVisibleText(para);
                            string num = "";

                            /* この段落のテキストが、直前に処理した図形（矢印や強調枠）の中のテキストと
                             * 完全に一致する場合は、Word上でその図形の陰に隠れて表示されない「重複した通常の段落」
                             * （コピー&ペースト等の残骸）とみなし、出力自体を丸ごとスキップする
                             */
                            if( a_outputImage_flg &&
                                !string.IsNullOrEmpty( text ) &&
                                recentShapeTexts.Contains( text.Trim() ) )
                            {
                                continue;
                            }

                            bool isHeading_flg = WordHelper.NumberingTypeEn.HEADING == numberingType;
                            bool isList_flg = WordHelper.NumberingTypeEn.LIST == numberingType;

                            /* この段落が新しい章の開始（新規ファイルを作るべき見出し）かどうか */
                            bool isNewChapter_flg = false;

                            /* この段落が、直前の章と重複する見出し（誤って紛れ込んだ空・重複見出し段落）であり、
                             * 出力自体を丸ごとスキップすべきかどうか
                             */
                            bool isDuplicateHeading_flg  = false;

                            if( isHeading_flg )
                            {
                                /* engine.Generateはカウンタをインクリメントするため、後で重複見出しと判明した場合に
                                 * 巻き戻せるよう、呼び出し前の状態を保存しておく
                                 */
                                Dictionary<int, int> engineStateBeforeGenerate = engine.SaveState();
                                bool usedEngineGenerate_flg = true;

                                if( numId.HasValue && numberingMap.ContainsKey( numId.Value ) )
                                {
                                    /* Wordの番号定義（numPr）から章番号を取得できる場合 */
                                    int levelValue = level ?? 0;
                                    num = engine.Generate( numberingMap[numId.Value], levelValue );
                                    usedEngineGenerate_flg = true;
                                }
                                
                                if( string.IsNullOrEmpty( num ) &&
                                    WordHelper.TryExtractLeadingNumber( text, out string extractedNum, out string extractedTitle ) )
                                {
                                    /* Wordの番号定義が無い場合、または番号定義はあってもレベルに対応する書式が
                                     * 取得できず番号を生成できなかった場合は、見出しテキスト先頭の数字パターン
                                     * （"3" "4.1"など）を章番号として代用する */
                                    num = extractedNum;
                                    text = extractedTitle;
                                }

                                /* ファイルを分割する条件はあくまで「章番号を取得できたかどうか」とする。
                                 * （本文が空でアウトラインレベルだけ設定された見出しや、番号を伴わない見出しスタイルの段落まで
                                 *   新しい章として分割してしまわないようにするため）
                                 */
                                if( !string.IsNullOrEmpty( num ) )
                                {
                                    /* 章番号は取得できたがタイトルが空の場合（見出しが章番号のみの場合）は、
                                     * 次の空白でない段落をタイトルとして補完する
                                     */
                                    bool isEmptyHeadingText_flg = string.IsNullOrEmpty( text );

                                    if( isEmptyHeadingText_flg )
                                    {
                                        string fallbackTitle = WordHelper.FindNextNonBlankParagraphText( elements, elementIndex + 1 );
                                        if( !string.IsNullOrEmpty( fallbackTitle ) )
                                        {
                                            text = fallbackTitle;
                                        }
                                    }

                                    isNewChapter_flg = !string.IsNullOrEmpty( text );

                                    /* 採用予定のタイトルが、直前に採用した章タイトルと完全に同一である場合は、
                                     * Word上に紛れ込んだ空・非表示の重複見出し段落（コピー&ペーストの残骸や
                                     * 削除履歴の残骸など）を誤って新章として分割してしまったものとみなし、
                                     * 新規ファイル作成を取り消す
                                     */
                                    if( isNewChapter_flg && text == lastChapterTitle )
                                    {
                                        isNewChapter_flg = false;
                                        isDuplicateHeading_flg = true;

                                        /* 章番号カウンタが実際には消費されなかったことにするため、
                                         * engine.Generate呼び出し前の状態に巻き戻す
                                         */
                                        if( usedEngineGenerate_flg )
                                        {
                                            engine.RestoreState( engineStateBeforeGenerate );
                                        }
                                    }

#if DEBUG
                                    onLogUpdate?.Invoke( $"[DEBUG] idx={elementIndex} numId={numId} level={level} num=\"{num}\" isEmptyHeadingText={isEmptyHeadingText_flg} text=\"{text}\" isNewChapter={isNewChapter_flg} lastChapterTitle=\"{lastChapterTitle}\"" );
#endif

                                    /* 重複見出しと判定した段落は、章番号・タイトルとして採用しないだけでなく、
                                     * 本文としても出力せずに丸ごとスキップする
                                     * （Word上に紛れ込んだ空・重複見出し段落の残骸を出力してしまわないようにするため）
                                     */
                                    if( isDuplicateHeading_flg )
                                    {
                                        continue;
                                    }

                                    if( isNewChapter_flg )
                                    {
                                        /* 現在の章番号を更新(同じ章の画像ファイルのファイル名に使用する) */
                                        currentNum = num;

                                        /* 今回採用した章タイトルを記憶する */
                                        lastChapterTitle = text;

                                        /* 章見出し自身を深さ1として、階層スタックを初期化する
                                         * （以降に出現する番号を持たない見出しの相対的な深さの基準にする）
                                         */
                                        headingStack.Clear();
                                        headingStack.Add( ( level ?? 0, 1 ) );
                                    }
                                }
                            }

                            /* 箇条書きの場合、Wordのレベルの書式に応じてMarkDown標準の記法を選択する
                             * （記号（Bullet）の場合は "-"、それ以外（連番・丸数字・アルファベットなど）の場合は "1." とする）
                             * 箇条書き番号が無効な場合は、書式に関わらず固定で "-" を使用する
                             */
                            string listMarkerFormat = "-";
                            if( isList_flg &&
                                a_outputListNumber_flg &&
                                numId.HasValue &&
                                numberingMap.ContainsKey( numId.Value ) )
                            {
                                int levelValue = level ?? 0;
                                NumberingDefinition def = numberingMap[numId.Value];

                                if( def.Levels.ContainsKey( levelValue ) &&
                                    Word.NumberFormatValues.Bullet != def.Levels[levelValue].format )
                                {
                                    /* Bullet以外（Decimal, DecimalEnclosedCircle, LowerLetterなど）は連番として扱う */
                                    listMarkerFormat = "1.";
                                }
                            }

                            /* 章番号を取得した場合は新規ファイルを作成する */
                            if( isNewChapter_flg )
                            {
                                /* 章番号を取得した場合 */

                                /* 章番号 章タイトル のファイルを追加 */
                                fileName = MarkDownWriter.SafeFileName( ( num + " " + text ).Trim() ) + ".md";
                                filePath = Path.Combine( a_outputDir, fileName );
                                md.NewFile( filePath );

                                /* ログにファイル名を表示 */
                                onLogUpdate( fileName );
                            }

                            /* Wordファイル「画像」の処理
                             * （画像を伴わない図形のみが複数の段落にまたがって離れた位置に配置されている場合、
                             *   それらをまとめて1枚の画像として合成することがあるため、消費した段落数の分だけ
                             *   ループのインデックスを進める）
                             */
                            if( a_outputImage_flg )
                            {
                                List<WordImageData> imageList = WordImageHelper.GetImages( doc.MainDocumentPart, elements, elementIndex, out int consumedCount );
                                elementIndex += consumedCount - 1;

                                foreach( WordImageData imageData in imageList )
                                {
                                    string imageFileName = m_imageFileNameGenerator.CreateFileName( currentNum, imageData.contentType );
                                    string imageDirectory = Path.Combine( a_outputDir, "images" );

                                    if( !Directory.Exists( imageDirectory ) )
                                    {
                                        Directory.CreateDirectory( imageDirectory );
                                    }

                                    string imagePath = Path.Combine( imageDirectory, imageFileName );
                                    File.WriteAllBytes( imagePath, imageData.imageData );

                                    /* MarkDown標準の画像記法（![alt](src)）だと、ビューアのテーマによっては
                                     * 画像だけ中央揃えで表示されてしまうことがあるため、
                                     * imgタグに明示的なスタイルを指定して左寄せを固定する
                                     */
                                    md.AddLine( "<img src=\"images/" + imageFileName + "\" alt=\"" + imageFileName + "\" style=\"display:block;margin:0;\" />" );
                                }
                            }

                            /* この段落の図形に設定されていたテキストを記録しておく（次の段落以降の重複判定に使用する）。
                             * 図形が無かった場合は、直前までの記録をクリアして重複判定の対象範囲を限定する
                             * （図形の直後に連続する段落だけを重複判定の対象とするため）
                             */
                            List<string> shapeTexts = a_outputImage_flg ? ShapeOverlayCompositor.GetShapeTexts( para ) : new List<string>();
                            if( 0 < shapeTexts.Count )
                            {
                                foreach( string shapeText in shapeTexts )
                                {
                                    recentShapeTexts.Add( shapeText );
                                }
                            }
                            else
                            {
                                recentShapeTexts.Clear();
                            }

                            /* 行出力（段落内改行(Shift+Enter)がある場合は複数行に分けて出力する） */
                            string[] textLines = text.Split( new[] { "\r\n", "\n" }, StringSplitOptions.None );

                            if( isList_flg )
                            {
                                /* 箇条書きの場合：MarkDown標準の記法（"-" または "1."）を使用する */
                                int indent = (level ?? 0) * 2;
                                string bulletPrefix = new string( ' ', indent ) + listMarkerFormat + " ";
                                string continuationIndent = new string( ' ', bulletPrefix.Length );

                                for( int i = 0; i < textLines.Length; i++ )
                                {
                                    /* 1行目は箇条書きの記号を付け、2行目以降は記号の位置に合わせてインデントする */
                                    string prefix = ( 0 == i ) ? bulletPrefix : continuationIndent;

                                    /* 最終行以外は、MarkDownの強制改行のため行末に半角スペースを2つ付与する */
                                    bool hasMoreLines_flg = ( i < textLines.Length - 1 );
                                    md.AddLine( prefix + textLines[i] + ( hasMoreLines_flg ? "  " : "" ) );
                                }
                            }
                            else
                            {
                                /* 見出しまたは通常の行の場合 */

                                /* 章番号を持つ見出し（新規ファイルの先頭行）は "#" 固定とする。
                                 * 章番号を持たず、同じファイル内に統合される見出しは、
                                 * 「直前に出現した見出しとの相対的な位置関係」から深さを決める（最大6段階）。
                                 * Wordのアウトラインレベルの数値差ではなく、実際に出現した見出しの並びを基準にする。
                                 */
                                string headingPrefix = "";
                                if( isNewChapter_flg )
                                {
                                    headingPrefix = "# ";
                                }
                                else if( isHeading_flg && !string.IsNullOrEmpty( text ) )
                                {
                                    int thisLevel = level ?? 0;

                                    /* 現在の見出しと同じか、それより深い階層として扱っていた要素をスタックから取り除く */
                                    while( 0 < headingStack.Count && thisLevel <= headingStack[headingStack.Count - 1].Level )
                                    {
                                        headingStack.RemoveAt( headingStack.Count - 1 );
                                    }

                                    /* 残った要素（直近の親にあたる見出し）の1つ下の深さとする */
                                    int parentDepth = ( 0 < headingStack.Count ) ? headingStack[headingStack.Count - 1].Depth : 1;
                                    int assignedDepth = parentDepth + 1;

                                    headingStack.Add( ( thisLevel, assignedDepth ) );

                                    int headingDepth = Math.Min( assignedDepth, 6 );
                                    headingPrefix = new string( '#', headingDepth ) + " ";
                                }

                                for( int i = 0; i < textLines.Length; i++ )
                                {
                                    /* 1行目のみ章番号・見出し記法を付与する */
                                    string line = ( 0 == i ) ? ( headingPrefix + $"{num} {textLines[i]}".Trim() ) : textLines[i];

                                    /* 最終行以外は、MarkDownの強制改行のため行末に半角スペースを2つ付与する */
                                    bool hasMoreLines_flg = ( i < textLines.Length - 1 );
                                    md.AddLine( line + ( hasMoreLines_flg ? "  " : "" ) );
                                }
                            }

                            md.AddLine( "" );
                            continue;
                        }

                        /* Wordファイル「表」の処理 */
                        Word.Table table = element as Word.Table;
                        if( null != table )
                        {
                            ConvertTable( doc.MainDocumentPart, table, md, a_outputDir, currentNum, a_outputImage_flg );

                            md.AddLine( "" );
                            continue;
                        }
                    }

                    /* 最後のファイルを保存する */
                    md.Save();
                }
            }
            catch( Exception ex )
            {
                string errMsg = "エラーが発生しました" + Environment.NewLine + Environment.NewLine + ex.Message;
                Console.WriteLine( errMsg );
                onLogUpdate( errMsg );
                System.Windows.MessageBox.Show( errMsg );
            }
            finally
            {
                if( !a_token.IsCancellationRequested )
                {
                    onProgressUpdate?.Invoke( 100 );
                }

                /* 一時ファイルを削除 */
                try
                {
                    System.IO.File.Delete( tempPath );
                }
                catch( Exception ex )
                {
                    Console.WriteLine( ex.Message );
                    onLogUpdate( ex.Message );
                }
            }
        }


        /// <summary>
        /// Word の表を MarkDown の表形式データへ変換し、指定された ファイル に追記する。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart（セル内画像の取得に使用）</param>
        /// <param name="a_table">変換元の Word の表</param>
        /// <param name="a_md">出力先 MarkDown ファイル</param>
        /// <param name="a_outputDir">画像ファイルの出力先ディレクトリ（MarkDownファイルの出力先）</param>
        /// <param name="a_headingNumber">画像ファイル名生成に使用する章番号</param>
        /// <param name="a_outputImage_flg">画像を出力するか否か</param>
        private void ConvertTable( MainDocumentPart a_mainDocumentPart, Word.Table a_table, MarkDownWriter a_md, string a_outputDir, string a_headingNumber, bool a_outputImage_flg )
        {
            bool headerDone = false;

            /* -----------------------------------------------------------------
             * Word 表の各行を順番に処理する
             * ----------------------------------------------------------------- */
            foreach( Word.TableRow tr in a_table.Elements<Word.TableRow>() )
            {
                List<string> cols = new List<string>();

                /* -------------------------------------------------------------
                 * Word 行内の各セルを順番に処理する
                 * ------------------------------------------------------------- */
                foreach( Word.TableCell tc in tr.Elements<Word.TableCell>() )
                {
                    string text = WordHelper.GetVisibleText(tc).Replace( "\n", " " );

                    /* セル内に画像が存在する場合は、画像ファイルを出力してimgタグをテキストに追加する
                     * （表のセルはMarkDown上1行で表現する必要があるため、改行を含む標準の画像記法ではなく
                     *   本文の画像出力と同じimgタグをテキストの末尾に連結する）
                     */
                    if( a_outputImage_flg )
                    {
                        string imageTags = ConvertCellImages( a_mainDocumentPart, tc, a_outputDir, a_headingNumber );

                        if( !string.IsNullOrEmpty( imageTags ) )
                        {
                            text = string.IsNullOrEmpty( text ) ? imageTags : ( text + " " + imageTags );
                        }
                    }

                    /* セルの GridSpan（横結合）を取得し、colspan に合わせてセルを展開する
                     * GridSpan が 1 の場合は通常どおり 1 列分を追加し、2 以上なら親セルの
                     * 表示内容を追加した後、残りを空セルで埋める。
                     */
                    var props = tc.GetFirstChild<Word.TableCellProperties>();
                    Word.GridSpan gridSpan = props?.GetFirstChild<Word.GridSpan>();
                    int span = ( gridSpan != null ) ? gridSpan.Val.Value : 1;

                    cols.Add( text );
                    for( int i = 1; i < span; i++ )
                    {
                        cols.Add( "" );
                    }
                }

                /* -------------------------------------------------------------
                 * 1 行分のデータを MarkDown に出力する
                 * ------------------------------------------------------------- */
                a_md.AddTableRow( cols.ToArray() );

                if( !headerDone )
                {
                    a_md.AddTableSeparator( cols.Count );
                    headerDone = true;
                }
            }
        }


        /// <summary>
        /// 表のセル内に存在する画像をファイルに保存し、参照用のimgタグ文字列を生成する。
        /// </summary>
        /// <param name="a_mainDocumentPart">MainDocumentPart</param>
        /// <param name="a_cell">対象セル</param>
        /// <param name="a_outputDir">画像ファイルの出力先ディレクトリ（MarkDownファイルの出力先）</param>
        /// <param name="a_headingNumber">画像ファイル名生成に使用する章番号</param>
        /// <returns>セル内の画像すべてに対するimgタグを連結した文字列（画像が無い場合は空文字列）</returns>
        private string ConvertCellImages( MainDocumentPart a_mainDocumentPart, Word.TableCell a_cell, string a_outputDir, string a_headingNumber )
        {
            StringBuilder sb = new StringBuilder();

            /* セル内の段落ごとに画像を取得する */
            foreach( Word.Paragraph para in a_cell.Elements<Word.Paragraph>() )
            {
                List<WordImageData> imageList = WordImageHelper.GetImages( a_mainDocumentPart, para );

                foreach( WordImageData imageData in imageList )
                {
                    string imageFileName = m_imageFileNameGenerator.CreateFileName( a_headingNumber, imageData.contentType );
                    string imageDirectory = Path.Combine( a_outputDir, "images" );

                    if( !Directory.Exists( imageDirectory ) )
                    {
                        Directory.CreateDirectory( imageDirectory );
                    }

                    string imagePath = Path.Combine( imageDirectory, imageFileName );
                    File.WriteAllBytes( imagePath, imageData.imageData );

                    if( 0 < sb.Length )
                    {
                        sb.Append( " " );
                    }

                    sb.Append( "<img src=\"images/" + imageFileName + "\" alt=\"" + imageFileName + "\" style=\"display:block;margin:0;\" />" );
                }
            }

            return sb.ToString();
        }


    }
}
