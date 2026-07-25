using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

                            bool isHeading_flg = WordHelper.NumberingTypeEn.HEADING == numberingType;
                            bool isList_flg = WordHelper.NumberingTypeEn.LIST == numberingType;

                            /* この段落が新しい章の開始（新規ファイルを作るべき見出し）かどうか */
                            bool isNewChapter_flg = false;

                            if( isHeading_flg )
                            {
                                if( numId.HasValue && numberingMap.ContainsKey( numId.Value ) )
                                {
                                    /* Wordの番号定義（numPr）から章番号を取得できる場合 */
                                    int levelValue = level ?? 0;
                                    num = engine.Generate( numberingMap[numId.Value], levelValue );
                                }
                                else if( WordHelper.TryExtractLeadingNumber( text, out string extractedNum, out string extractedTitle ) )
                                {
                                    /* Wordの番号定義が無い場合は、見出しテキスト先頭の数字パターン（"3" "4.1"など）を章番号として代用する */
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
                                    if( string.IsNullOrEmpty( text ) )
                                    {
                                        string fallbackTitle = WordHelper.FindNextNonBlankParagraphText( elements, elementIndex + 1 );
                                        if( !string.IsNullOrEmpty( fallbackTitle ) )
                                        {
                                            text = fallbackTitle;
                                        }
                                    }

                                    isNewChapter_flg = !string.IsNullOrEmpty( text );

                                    if( isNewChapter_flg )
                                    {
                                        /* 現在の章番号を更新(同じ章の画像ファイルのファイル名に使用する) */
                                        currentNum = num;
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

                            /* Wordファイル「画像」の処理 */
                            if( a_outputImage_flg )
                            {
                                List<WordImageData> imageList = WordImageHelper.GetImages( doc.MainDocumentPart, para );

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

                                    md.AddLine( "![" + imageFileName + "](images/" + imageFileName + ")" );
                                }
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

                                /* 章の見出しとして扱われた行は、MarkDownの見出し記法 "# " を先頭に付与する */
                                /* このアプリでは章毎に別のファイルに書き出すので階層の深さに関係なく#の数を1つに固定しています */
                                bool isHeadingRow_flg = isNewChapter_flg;
                                string headingPrefix = "";
                                if( isHeadingRow_flg )
                                {
                                    headingPrefix = "# ";
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
                            ConvertTable( table, md );

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
        /// <param name="a_table">変換元の Word の表</param>
        /// <param name="a_md">出力先 MarkDown ファイル</param>
        private void ConvertTable( Word.Table a_table, MarkDownWriter a_md )
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
                    string text = WordHelper.GetVisibleText(tc);
                    cols.Add( text.Replace( "\n", " " ) );
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


    }
}
