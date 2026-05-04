using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

                        /* Excelの書式設定を生成 */
                        ExcelHelper.CreateStylesheet( workbookPart );

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
                                ExcelHelper.SetRow( sheetData, row++, new List<CellData>() { numData, textData } );

                                continue;
                            }

                            /* Wordファイル「表」の処理 */
                            Word.Table table = element as Word.Table;
                            if( null != table )
                            {
                                /* Excelワークシートを追加 */
                                if( null == wsPart )
                                {
                                    wsPart = ExcelHelper.CreateWorksheet( workbookPart, sheets, "Sheet1", sheetId++, out sheetData );
                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }

                                foreach( Word.TableRow tr in table.Elements<Word.TableRow>() )
                                {
                                    List<CellData> values = new List<CellData>();

                                    /* 先頭列は章番号に使用するので1列ずらす */
                                    values.Add( new CellData() { Value = "" } );

                                    foreach( Word.TableCell tc in tr.Elements<Word.TableCell>() )
                                    {
                                        /* Word「表」プロパティを取得 */
                                        Word.TableCellProperties props = tc.TableCellProperties;

                                        /* Wordセル罫線有無 */
                                        Word.TableCellBorders borders = props?.GetFirstChild<Word.TableCellBorders>();
                                        bool hasBorder = ( null != borders );

                                        /* セルデータをセット */
                                        values.Add( new CellData() { Value = WordHelper.GetVisibleText( tc ), BorderTop = true, BorderBottom = true, BorderLeft = true, BorderRight = true } );

                                        /* GridSpan */
                                        Word.GridSpan gridSpan = props?.GetFirstChild<Word.GridSpan>();
                                        int span = ( null != gridSpan ) ? gridSpan.Val.Value : 1;

                                        /* 横結合セル分の空セル追加 */
                                        for( int i = 1; i < span; i++ )
                                        {
                                            values.Add( new CellData() { Value = "", BorderTop = true, BorderBottom = true, BorderLeft = true, BorderRight = true } );
                                        }
                                    }

                                    /* 行出力 */
                                    ExcelHelper.SetRow( sheetData, row++, values );
                                }

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

    }
}
