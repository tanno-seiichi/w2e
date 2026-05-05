using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using w2e.excel;
using w2e.file;
using w2e.word;
using Excel = DocumentFormat.OpenXml.Spreadsheet;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace w2e.converter
{
    public class W2EConverter
    {
        private const int PROGRESS_MIN_VALUE = 1;
        private const int PROGRESS_WORD_RANGE = 40;
        private const int PROGRESS_EXCEL_RANGE = 60;
        
        public delegate void UpdateProgressDelegate( int a_value );
        public delegate void UpdateLogDelegate( string a_value );
        
        public static UpdateProgressDelegate onProgressUpdate { get; set; }

        public static UpdateLogDelegate onLogUpdate { get; set; }

        public static void Convert( string wordPath, string excelPath, CancellationToken a_token )
        {
            onProgressUpdate?.Invoke( PROGRESS_MIN_VALUE );
            string tempPath = FileCopy.CreateTempCopy( wordPath );

            try
            {
                /* Wordファイルを読込 */
                using( WordprocessingDocument doc = WordprocessingDocument.Open( tempPath, false ) )
                {
                    Word.Body body = doc.MainDocumentPart.Document.Body;
                    StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;

                    Dictionary<int, NumberingDefinition> numberingMap = WordHelper.LoadNumbring( doc );
                    NumberingEngine engine = new NumberingEngine();

                    /* Excelファイルを生成 */
                    using( SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create( excelPath, SpreadsheetDocumentType.Workbook ) )
                    {
                        /* Excelワークブックを生成 */
                        WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                        workbookPart.Workbook = new Excel.Workbook();

#if BORDER_1
                        /* Excelの書式設定を生成 */
                        ExcelHelper.CreateStylesheet( workbookPart );
#else
                        /* Excelのスタイルシートを初期化 */
                        ExcelHelper.InitializeStylesheet( workbookPart );

                        /* Excelのスタイルシートに登録済のスタイルを再利用するためのキャッシュを生成 */
                        var cache = new Dictionary<string, uint>();
#endif

                        /* Excelワークブックのシートを追加する準備 */
                        Excel.Sheets sheets = workbookPart.Workbook.AppendChild( new Excel.Sheets() );

                        WorksheetPart wsPart = null;
                        Excel.SheetData sheetData = null;
                        uint sheetId = 1;
                        int row = 1;

                        int total = body.Elements().Count();
                        int current = 0;

                        /* プログレスバーをWordファイル読込終了まで進める */
                        onProgressUpdate?.Invoke( PROGRESS_WORD_RANGE );

                        /* Wordファイルの要素ごとに処理 */
                        foreach( OpenXmlElement element in body.Elements() )
                        {
                            /* 処理中断が要求されていたらループを抜ける */
                            if( a_token.IsCancellationRequested ) { break; }

                            /* プログレスバーを更新 */
                            current++;
                            int progress = PROGRESS_WORD_RANGE + (int)(current * PROGRESS_EXCEL_RANGE / total );
                            onProgressUpdate?.Invoke( progress );

                            /* Wordファイル「段落」の処理 */
                            Word.Paragraph para = element as Word.Paragraph;
                            if( null != para )
                            {
                                var info = WordHelper.GetNumberingInfo( para, stylePart );
                                int? numId = info.numId;
                                int? level = info.level;

                                CellData textData = new CellData() { Value = WordHelper.GetVisibleText( para ) };
                                CellData numData = new CellData() { Value = "" };

                                if( numId.HasValue &&
                                    numberingMap.ContainsKey( numId.Value ) )
                                {
                                    int levelValue = level.HasValue ? level.Value : 0;
                                    numData.Value = engine.Generate( numberingMap[numId.Value], levelValue );
                                }

                                /* シートが未登録の場合、または章番号を取得した場合は新規シートを追加する */
                                if( null == wsPart )
                                {
                                    /* シートが未登録の場合 */
    
                                    /* 先頭シートを追加 */
                                    wsPart = ExcelHelper.CreateWorksheet( workbookPart, sheets, "トップ", sheetId++, out sheetData );
                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }
                                else if( !string.IsNullOrEmpty( numData.Value ) )
                                {
                                    /* 章番号を取得した場合 */
    
                                    /* 章番号 章タイトル のシートを追加 */
                                    wsPart = ExcelHelper.CreateWorksheet( workbookPart, sheets, ExcelHelper.SafeSheetName( numData.Value + " " + textData.Value ), sheetId++, out sheetData );
                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }

                                /* 行出力 */
#if BORDER_1
                                ExcelHelper.SetRow( sheetData, row++, new List<CellData>() { numData, textData } );
#else
                                ExcelHelper.SetRow( workbookPart, sheetData, row++, new List<CellData>() { numData, textData }, cache );
#endif
                                continue;
                            }

                            /* Wordファイル「表」の処理 */
                            Word.Table table = element as Word.Table;
                            if( null != table )
                            {
                                /* Excelワークシートを追加 */
                                if( null == wsPart )
                                {
                                    /* シートが未登録の場合 */

                                    wsPart = ExcelHelper.CreateWorksheet( workbookPart, sheets, "トップ", sheetId++, out sheetData );
                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }

#if BORDER_1
                                ConvertTable( table, sheetData, ref row );
#else
                                ConvertTable( workbookPart, table, sheetData, ref row, cache );
#endif
                                row++;
                                continue;
                            }
                        }
                    }
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

#if BORDER_1
        /// <summary>
        /// Word の表を Excel の表形式データへ変換し、指定された SheetData に追記する。
        /// </summary>
        /// <param name="table">変換元の Word の表</param>
        /// <param name="sheetData">出力先 Excel シートデータ</param>
        /// <param name="row">Excel の出力開始行番号（出力後は次の行番号へ更新される）</param>
        private static void ConvertTable( Word.Table table, SheetData sheetData, ref int row )
        {
            /* -----------------------------------------------------------------
             * Word 表の各行を順番に処理する
             * ----------------------------------------------------------------- */
            foreach( Word.TableRow tr in table.Elements<Word.TableRow>() )
            {
                /* -------------------------------------------------------------
                 * 1 行分の Excel セルデータを格納するリスト
                 * この values がそのまま Excel の 1 行になる
                 * ------------------------------------------------------------- */
                List<CellData> values = new List<CellData>();

                /* -------------------------------------------------------------
                 * 先頭列は章番号など別用途で使用するため、
                 * Word 表の内容は 1 列右にずらして出力する
                 * ------------------------------------------------------------- */
                values.Add( new CellData() { Value = "" } );

                /* -------------------------------------------------------------
                 * Word 行内の各セルを順番に処理する
                 * ------------------------------------------------------------- */
                foreach( Word.TableCell tc in tr.Elements<Word.TableCell>() )
                {
                    /* ---------------------------------------------------------
                     * セルのプロパティを取得
                     * （結合情報、罫線情報などが含まれる）
                     * --------------------------------------------------------- */
                    Word.TableCellProperties props = tc.TableCellProperties;

                    /* ---------------------------------------------------------
                     * Word セルの罫線情報を取得
                     * 現状この情報は未使用
                     * （必要なら Excel 側の罫線反映に利用可能）
                     * --------------------------------------------------------- */
                    Word.TableCellBorders borders = props?.GetFirstChild<Word.TableCellBorders>();

                    /* ---------------------------------------------------------
                     * GridSpan を取得
                     *
                     * GridSpan は「このセルが横方向に何列分を占有しているか」
                     * を表す。
                     *
                     * 例:
                     *   GridSpan = 3 の場合、
                     *   このセルは Excel 上で 3 列分に相当する。
                     * --------------------------------------------------------- */
                    Word.GridSpan gridSpan = props?.GetFirstChild<Word.GridSpan>();
                    int span = ( null != gridSpan ) ? gridSpan.Val.Value : 1;

                    /* ---------------------------------------------------------
                     * セルに表示されている文字列を取得して Excel セルへ設定
                     *
                     * 現在は Word 側の罫線有無に関係なく、
                     * すべての辺に罫線ありとして出力している
                     * --------------------------------------------------------- */
                    values.Add(
                        new CellData()
                        {
                            Value = WordHelper.GetVisibleText( tc ),
                            BorderTop = true,
                            BorderBottom = true,
                            BorderLeft = true,
                            BorderRight = ( 0 < span ) ? false : true
                        } );

                    /* ---------------------------------------------------------
                     * 横方向に結合されている残りの列数分、
                     * Excel 側で位置合わせ用の空セルを追加する
                     *
                     * 先頭セルはすでに追加済みなので i = 1 から開始
                     * --------------------------------------------------------- */
                    for( int i = 1; i < span; i++ )
                    {
                        values.Add(
                            new CellData()
                            {
                                Value = "",
                                BorderTop = true,
                                BorderBottom = true,
                                BorderLeft = false,
                                BorderRight = true
                            } );
                    }
                }

                /* -------------------------------------------------------------
                 * 1 行分のセルデータを Excel に出力する
                 * 出力後、次の行番号へ進める
                 * ------------------------------------------------------------- */
                ExcelHelper.SetRow( sheetData, row++, values );
            }
        }
#else
        /// <summary>
        /// Word の表を Excel の表形式データへ変換し、指定された SheetData に追記する。
        /// </summary>
        /// <param name="table">変換元の Word の表</param>
        /// <param name="sheetData">出力先 Excel シートデータ</param>
        /// <param name="row">Excel の出力開始行番号（出力後は次の行番号へ更新される）</param>
        private static void ConvertTable( WorkbookPart a_wbPart, Word.Table table, SheetData sheetData, ref int row, Dictionary<string, uint> a_cache )
        {
            /* -----------------------------------------------------------------
             * Word 表の各行を順番に処理する
             * ----------------------------------------------------------------- */
            foreach( Word.TableRow tr in table.Elements<Word.TableRow>() )
            {
                /* -------------------------------------------------------------
                 * 1 行分の Excel セルデータを格納するリスト
                 * この values がそのまま Excel の 1 行になる
                 * ------------------------------------------------------------- */
                List<CellData> values = new List<CellData>();

                /* -------------------------------------------------------------
                 * 先頭列は章番号など別用途で使用するため、
                 * Word 表の内容は 1 列右にずらして出力する
                 * ------------------------------------------------------------- */
                values.Add( new CellData() { Value = "" } );

                /* -------------------------------------------------------------
                 * Word 行内の各セルを順番に処理する
                 * ------------------------------------------------------------- */
                foreach( Word.TableCell tc in tr.Elements<Word.TableCell>() )
                {
                    /* ---------------------------------------------------------
                     * セルのプロパティを取得
                     * （結合情報、罫線情報などが含まれる）
                     * --------------------------------------------------------- */
                    Word.TableCellProperties props = tc.TableCellProperties;

                    /* ---------------------------------------------------------
                     * Word セルの罫線情報を取得
                     * 現状この情報は未使用
                     * （必要なら Excel 側の罫線反映に利用可能）
                     * --------------------------------------------------------- */
                    Word.TableCellBorders borders = props?.GetFirstChild<Word.TableCellBorders>();

                    /* ---------------------------------------------------------
                     * GridSpan を取得
                     *
                     * GridSpan は「このセルが横方向に何列分を占有しているか」
                     * を表す。
                     *
                     * 例:
                     *   GridSpan = 3 の場合、
                     *   このセルは Excel 上で 3 列分に相当する。
                     * --------------------------------------------------------- */
                    Word.GridSpan gridSpan = props?.GetFirstChild<Word.GridSpan>();
                    int span = ( null != gridSpan ) ? gridSpan.Val.Value : 1;

                    /* ---------------------------------------------------------
                     * セルに表示されている文字列を取得して Excel セルへ設定
                     *
                     * 現在は Word 側の罫線有無に関係なく、
                     * すべての辺に罫線ありとして出力している
                     * --------------------------------------------------------- */
                    values.Add(
                        new CellData()
                        {
                            Value = WordHelper.GetVisibleText( tc ),
                            BorderTop = true,
                            BorderBottom = true,
                            BorderLeft = true,
                            BorderRight = ( 1 < span ) ? false : true
                        } );

                    /* ---------------------------------------------------------
                     * 横方向に結合されている残りの列数分、
                     * Excel 側で位置合わせ用の空セルを追加する
                     *
                     * 先頭セルはすでに追加済みなので i = 1 から開始
                     * --------------------------------------------------------- */
                    for( int i = 1; i < span; i++ )
                    {
                        values.Add(
                            new CellData()
                            {
                                Value = "",
                                BorderTop = true,
                                BorderBottom = true,
                                BorderLeft = false,
                                BorderRight = true
                            } );
                    }
                }

                /* -------------------------------------------------------------
                 * 1 行分のセルデータを Excel に出力する
                 * 出力後、次の行番号へ進める
                 * ------------------------------------------------------------- */
                ExcelHelper.SetRow( a_wbPart, sheetData, row++, values, a_cache );
            }
        }
#endif
    }
}
