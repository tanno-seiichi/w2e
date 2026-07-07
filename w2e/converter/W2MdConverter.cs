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
        /// <param name="a_token">処理中断通知</param>
        public void Convert( string a_wordPath, string a_outputDir, bool a_outputImage_flg, CancellationToken a_token )
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

                    /* Wordファイルの要素ごとに処理 */
                    foreach( OpenXmlElement element in body.Elements() )
                    {
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

                            /* 有効な番号付情報と章タイトルの組合せを検出したら章番号を設定する */
                            if( isHeading_flg &&
                                !string.IsNullOrEmpty( text ) &&
                                numId.HasValue &&
                                numberingMap.ContainsKey( numId.Value ) )
                            {
                                int levelValue = level ?? 0;
                                num = engine.Generate( numberingMap[numId.Value], levelValue );

                                /* 現在の章番号を更新(同じ章の画像ファイルのファイル名に使用する) */
                                currentNum = num;
                            }

                            /* 章番号を取得した場合は新規ファイルを作成する */
                            if( !string.IsNullOrEmpty( num ) )
                            {
                                /* 章番号を取得した場合 */

                                /* 章番号 章タイトル のファイルを追加 */
                                fileName = MarkDownWriter.SafeFileName( num + " " + text ) + ".md";
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

                            /* 行出力 */
                            if( isList_flg )
                            {
                                /* 箇条書きの場合 */
                                int indent = (level ?? 0) * 2;
                                md.AddLine( new string( ' ', indent ) + "- " + text );
                            }
                            else
                            {
                                /* 見出しまたは通常の行の場合 */
                                md.AddLine( $"{num} {text}".Trim() );
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
