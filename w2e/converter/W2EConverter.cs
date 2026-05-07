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
    /// <summary>
    /// Wordファイルの内容をExcelファイルに書き出すクラス
    /// </summary>
    public class W2EConverter
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
        /// Excel書出完了時の進捗値を表す定数
        /// </summary>
        private const int PROGRESS_EXCEL_RANGE = 60;

        /// <summary>
        /// 先頭シートのシート名
        /// </summary>
        private const string TOP_SHEET_NAME = "トップ";

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
        /// Word → Excel の変換処理を実行する
        /// </summary>
        /// <param name="a_wordPath">Wordファイルのパス</param>
        /// <param name="a_excelPath">Excelファイルのパス</param>
        /// <param name="a_token">処理中断通知</param>
        public static void Convert( string a_wordPath, string a_excelPath, CancellationToken a_token )
        {
            onProgressUpdate?.Invoke( PROGRESS_MIN_VALUE );
            string tempPath = FileCopy.CreateTempCopy( a_wordPath );

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
                    using( SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create( a_excelPath, SpreadsheetDocumentType.Workbook ) )
                    {
                        /* Excelワークブックを生成 */
                        WorkbookPart wbPart = spreadsheet.AddWorkbookPart();
                        wbPart.Workbook = new Excel.Workbook();

                        /* Excelのスタイルシートを初期化 */
                        ExcelHelper.InitializeStylesheet( wbPart );

                        /* Excelのスタイルシートに登録済のスタイルを再利用するためのキャッシュを生成 */
                        var cache = new Dictionary<string, uint>();

                        /* Excelワークブックのシートを追加する準備 */
                        Excel.Sheets sheets = wbPart.Workbook.AppendChild( new Excel.Sheets() );

                        WorksheetPart wsPart = null;
                        Excel.SheetData sheetData = null;
                        string sheetName = TOP_SHEET_NAME;
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

                                CellData textData = new CellData() { text = WordHelper.GetVisibleText( para ) };
                                CellData numData = new CellData() { text = "" };

                                /* 有効な番号付情報と章タイトルの組合せを検出したら章番号を設定する */
                                if( !string.IsNullOrEmpty( textData.text ) &&
                                    numId.HasValue &&
                                    numberingMap.ContainsKey( numId.Value ) )
                                {
                                    int levelValue = level.HasValue ? level.Value : 0;
                                    numData.text = engine.Generate( numberingMap[numId.Value], levelValue );
                                }

                                /* シートが未登録の場合、または章番号を取得した場合は新規シートを追加する */
                                if( null == wsPart )
                                {
                                    /* シートが未登録の場合 */

                                    /* 先頭シートを追加 */
                                    wsPart = ExcelHelper.CreateWorksheet( wbPart, sheets, sheetName, sheetId++, out sheetData );

                                    /* ログにシート名を表示 */
                                    onLogUpdate( sheetName );
                                }
                                else if( !string.IsNullOrEmpty( numData.text ) )
                                {
                                    /* 章番号を取得した場合 */

                                    /* 章番号 章タイトル のシートを追加 */
                                    sheetName = ExcelHelper.SafeSheetName( numData.text + " " + textData.text );
                                    wsPart = ExcelHelper.CreateWorksheet( wbPart, sheets, sheetName, sheetId++, out sheetData );

                                    /* ログにシート名を表示 */
                                    onLogUpdate( sheetName );

                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }

                                /* 行出力 */
                                ExcelHelper.SetRow( wbPart, sheetData, row++, new List<CellData>() { numData, textData }, cache );
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

                                    /* 先頭シートを追加 */
                                    wsPart = ExcelHelper.CreateWorksheet( wbPart, sheets, sheetName, sheetId++, out sheetData );

                                    /* ログにシート名を表示 */
                                    onLogUpdate( sheetName );
                                }

                                ConvertTable( wbPart, table, sheetData, ref row, cache );

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


        /// <summary>
        /// Word の表を Excel の表形式データへ変換し、指定された SheetData に追記する。
        /// </summary>
        /// <param name="a_table">変換元の Word の表</param>
        /// <param name="a_sheetData">出力先 Excel シートデータ</param>
        /// <param name="a_row">Excel の出力開始行番号（出力後は次の行番号へ更新される）</param>
        private static void ConvertTable( WorkbookPart a_wbPart, Word.Table a_table, SheetData a_sheetData, ref int a_row, Dictionary<string, uint> a_cache )
        {
            /* -----------------------------------------------------------------
             * Word 表の全行を List にして index で参照できるようにする
             * ----------------------------------------------------------------- */
            List<Word.TableRow> rows = a_table.Elements<Word.TableRow>().ToList();

            /* -----------------------------------------------------------------
             * Word 表の各行を順番に処理する
             * ----------------------------------------------------------------- */
            for( int rowIndex = 0; rowIndex < rows.Count; rowIndex++ )
            {
                Word.TableRow tr = rows[rowIndex];

                /* -------------------------------------------------------------
                 * 1 行分の Excel セルデータを格納するリスト
                 * この values がそのまま Excel の 1 行になる
                 * ------------------------------------------------------------- */
                List<CellData> values = new List<CellData>();

                /* -------------------------------------------------------------
                 * 先頭列は章番号など別用途で使用するため、
                 * Word 表の内容は 1 列右にずらして出力する
                 * ------------------------------------------------------------- */
                values.Add( new CellData() { text = "" } );

                /* -------------------------------------------------------------
                 * この行のセル一覧を取得（列 index 用）
                 * ------------------------------------------------------------- */
                List<Word.TableCell> cells = tr.Elements<Word.TableCell>().ToList();

                /* -------------------------------------------------------------
                 * Word 行内の各セルを順番に処理する
                 * ------------------------------------------------------------- */
                for( int colIndex = 0; colIndex < cells.Count; colIndex++ )
                {
                    Word.TableCell tc = cells[colIndex];

                    /* ---------------------------------------------------------
                     * セルのプロパティを取得
                     * （結合情報、罫線情報などが含まれる）
                     * --------------------------------------------------------- */
                    Word.TableCellProperties props = tc.TableCellProperties;

                    /* ---------------------------------------------------------
                     * GridSpan（横結合の列数）を取得
                     * --------------------------------------------------------- */
                    Word.GridSpan gridSpan = props?.GetFirstChild<Word.GridSpan>();
                    int span = ( null != gridSpan ) ? gridSpan.Val.Value : 1;

                    /* ---------------------------------------------------------
                     * VerticalMerge を取得
                     *
                     * VerticalMerge.Val
                     * ・MergedCellValues.Restart : 縦結合セルの開始セル
                     * ・MergedCellValues.Continue : 縦結合セルの継続セル
                     * --------------------------------------------------------- */
                    Word.VerticalMerge vertical = props?.GetFirstChild<Word.VerticalMerge>();
                    bool isRestart_flg = false;
                    bool isContinue_flg = false;
                    if( null != vertical )
                    {
                        if( null == vertical.Val )
                        {
                            /* 結合セル */
                            isContinue_flg = true;
                        }
                        else
                        {
                            if( Word.MergedCellValues.Restart == vertical.Val.Value )
                            {
                                /* 開始セル */
                                isRestart_flg = true;
                            }
                        }
                    }

                    /* ---------------------------------------------------------
                     * 次の行の同じ列のセルで縦結合が継続しているか判定
                     * --------------------------------------------------------- */
                    bool hasNextVerticalMerge_flg = false;

                    if( isContinue_flg &&
                        rowIndex + 1 < rows.Count )
                    {
                        Word.TableRow nextRow = rows[rowIndex + 1];
                        List<Word.TableCell> nextCells = nextRow.Elements<Word.TableCell>().ToList();

                        if( colIndex < nextCells.Count )
                        {
                            var nextVmerge = nextCells[colIndex].TableCellProperties?.GetFirstChild<Word.VerticalMerge>();
                            hasNextVerticalMerge_flg = ( null != nextVmerge );
                        }
                    }

                    /* ---------------------------------------------------------
                     * 枠線（下）判定
                     * ・非縦結合            ： あり
                     * ・Restart             ： なし
                     * ・Continue + 次もあり ： なし（縦結合セルの中間）
                     * ・Continue + 次はなし ： あり（縦結合セルの末尾）
                     * --------------------------------------------------------- */
                    bool bottomBorder_flg = ( null == vertical ) ? true : ( isContinue_flg ? !hasNextVerticalMerge_flg : false );

                    /* ---------------------------------------------------------
                     * セルデータを追加（先頭セル）
                     * --------------------------------------------------------- */
                    values.Add(
                        new CellData()
                        {
                            text = isContinue_flg ? "" : WordHelper.GetCellText( tc ),
                            topBorder = ( null == vertical ) ? true : isRestart_flg,
                            bottomBorder = bottomBorder_flg,
                            leftBorder = true,
                            rightBorder = ( 1 < span ) ? false : true
                        } );

                    /* ---------------------------------------------------------
                     * 横方向に結合されている残りの列数分、
                     * Excel 側で位置合わせ用の空セルを追加する
                     * 結合セルの末尾まで枠線（右）は設定しない
                     *
                     * 先頭セルはすでに追加済みなので i = 1 から開始
                     * --------------------------------------------------------- */
                    for( int i = 1; i < span; i++ )
                    {
                        values.Add(
                            new CellData()
                            {
                                text = "",
                                topBorder = ( null == vertical ) ? true : isRestart_flg,
                                bottomBorder = bottomBorder_flg,
                                leftBorder = false,
                                rightBorder = ( i == span - 1 ) ? true : false
                            } );
                    }
                }

                /* -------------------------------------------------------------
                 * 1 行分のセルデータを Excel に出力する
                 * 出力後、次の行番号へ進める
                 * ------------------------------------------------------------- */
                ExcelHelper.SetRow( a_wbPart, a_sheetData, a_row++, values, a_cache );
            }
        }


    }
}
