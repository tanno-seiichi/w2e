using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using w2e.excel;
using w2e.file;
using w2e.markdown;
using w2e.word;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace w2e.converter
{
    /// <summary>
    /// Wordファイルの内容をMarkDownファイルに書き出すクラス
    /// </summary>
    public class W2MdConverter
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
        /// 進捗情報の処理を委譲するDelegate
        /// </summary>
        /// <param name="a_value"></param>
        public delegate void UpdateProgressDelegate( int a_value );

        /// <summary>
        /// ログ出力処理をを委譲するDelegate
        /// </summary>
        /// <param name="a_value"></param>
        public delegate void UpdateLogDelegate( string a_value );

        /// <summary>
        /// 進捗情報が更新された時の処理
        /// </summary>
        public static UpdateProgressDelegate onProgressUpdate { get; set; }

        /// <summary>
        /// ログが出力された時の処理
        /// </summary>
        public static UpdateLogDelegate onLogUpdate { get; set; }


        /// <summary>
        /// Word → MarkDown の変換処理を実行する
        /// </summary>
        /// <param name="a_wordPath">Wordファイルのパス</param>
        /// <param name="a_outputDir">MarkDownファイルの出力先ディレクトリ</param>
        /// <param name="a_token">処理中断通知</param>
        public static void Convert( string a_wordPath, string a_outputDir, CancellationToken a_token )
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

                    /* 現在のファイル情報を初期化 */
                    string currentFile = null;

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

                            string text = WordHelper.GetVisibleText(para);
                            string num = "";

                            /* 有効な番号付情報と章タイトルの組合せを検出したら章番号を設定する */
                            if( !string.IsNullOrEmpty( text ) &&
                                numId.HasValue &&
                                numberingMap.ContainsKey( numId.Value ) )
                            {
                                int levelValue = level ?? 0;
                                num = engine.Generate( numberingMap[numId.Value], levelValue );
                            }

                            /* ファイルが未登録の場合、または章番号を取得した場合は新規ファイルを作成する */
                            if( null == currentFile )
                            {
                                /* ファイルが未登録の場合 */

                                /* 先頭ファイルを追加 */
                                currentFile = Path.Combine( a_outputDir, "トップ.md" );
                                md.NewFile( currentFile );
                            }
                            else if( !string.IsNullOrEmpty( num ) )
                            {
                                /* 章番号を取得した場合 */

                                /* 章番号 章タイトル のファイルを追加 */
                                string fileName = ExcelHelper.SafeSheetName(num + " " + text) + ".md";
                                currentFile = Path.Combine( a_outputDir, fileName );
                                md.NewFile( currentFile );
                            }

                            /* 行出力 */
                            md.AddLine( $"{num} {text}".Trim() );
                            md.AddLine( "" );
                            continue;
                        }

                        /* Wordファイル「表」の処理 */
                        Word.Table table = element as Word.Table;
                        if( null != table )
                        {
                            /* ファイルが未登録の場合は新規ファイルを作成する */
                            if( null == currentFile )
                            {
                                /* ファイルが未登録の場合 */

                                currentFile = Path.Combine( a_outputDir, "トップ.md" );
                                md.NewFile( currentFile );
                            }

                            ConvertTable( table, md );

                            md.AddLine( "" );
                            continue;
                        }
                    }

                    /* 最後のファイルを保存する */
                    md.Save();
                }
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
        private static void ConvertTable( Word.Table a_table, MarkDownWriter a_md )
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
