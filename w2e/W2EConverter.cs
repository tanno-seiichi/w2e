using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Excel = DocumentFormat.OpenXml.Spreadsheet;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace w2e
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
            string tempPath = CreateTempCopy( wordPath );

            try
            {
                /* Wordファイルを読込 */
                using( WordprocessingDocument doc = WordprocessingDocument.Open( tempPath, false ) )
                {
                    Word.Body body = doc.MainDocumentPart.Document.Body;
                    StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;

                    Dictionary<int, NumberingDefinition> numberingMap = LoadNumbring( doc );
                    NumberingEngine engine = new NumberingEngine();

                    /* Excelファイルを生成 */
                    using( SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create( excelPath, SpreadsheetDocumentType.Workbook ) )
                    {
                        /* Excelワークブックを生成 */
                        WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                        workbookPart.Workbook = new Excel.Workbook();

                        /* Excelの書式設定を生成 */
                        CreateStylesheet( workbookPart );

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
                                var info = GetNumberingInfo( para, stylePart );
                                int? numId = info.numId;
                                int? level = info.level;

                                CellData textData = new CellData() { Value = GetVisibleText( para ) };
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
                                    wsPart = CreateWorksheet( workbookPart, sheets, "トップ", sheetId++, out sheetData );
                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }
                                else if( !string.IsNullOrEmpty( numData.Value ) )
                                {
                                    /* 章番号を取得した場合 */
    
                                    /* 章番号 章タイトル のシートを追加 */
                                    wsPart = CreateWorksheet( workbookPart, sheets, SafeSheetName( numData.Value + " " + textData.Value ), sheetId++, out sheetData );
                                    /* シートが変わったので行を先頭に戻す */
                                    row = 1;
                                }

                                /* 行出力 */
                                SetRow( sheetData, row++, new List<CellData>() { numData, textData } );

                                continue;
                            }

                            /* Wordファイル「表」の処理 */
                            Word.Table table = element as Word.Table;
                            if( null != table )
                            {
                                /* Excelワークシートを追加 */
                                if( null == wsPart )
                                {
                                    wsPart = CreateWorksheet( workbookPart, sheets, "Sheet1", sheetId++, out sheetData );
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
                                        values.Add( new CellData() { Value = GetVisibleText( tc ), BorderTop = true, BorderBottom = true, BorderLeft = true, BorderRight = true } );

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
                                    SetRow( sheetData, row++, values );
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

        private static string CreateTempCopy( string a_wordPath )
        {
            string dir = System.IO.Path.GetDirectoryName( a_wordPath );
            string tempPath = System.IO.Path.Combine( dir, System.IO.Path.GetFileNameWithoutExtension( a_wordPath ) + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + System.IO.Path.GetExtension( a_wordPath ) );
            System.IO.File.Copy( a_wordPath, tempPath, false );
            return tempPath;
        }

        #region Excel Helper

        class CellData
        {
            public string Value = "";
            public bool BorderTop = false;
            public bool BorderBottom = false;
            public bool BorderLeft = false;
            public bool BorderRight = false;
        }

        private static WorksheetPart CreateWorksheet( WorkbookPart wbPart, Excel.Sheets sheets, string sheetName, uint sheetId, out Excel.SheetData sheetData )
        {
            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();

            sheetData = new Excel.SheetData();
            wsPart.Worksheet = new Excel.Worksheet( sheetData );

            Excel.Sheet sheet = new Excel.Sheet();
            sheet.Id = wbPart.GetIdOfPart( wsPart );
            sheet.SheetId = sheetId;
            sheet.Name = sheetName;

            sheets.Append( sheet );
            return wsPart;
        }

        private static void CreateStylesheet( WorkbookPart workbookPart )
        {

            Excel.Fonts fonts = new Excel.Fonts( new Excel.Font() );
            Excel.Fills fills = new Excel.Fills( new Excel.Fill( new Excel.PatternFill() ) );

            Excel.Borders borders = new Excel.Borders();
            borders.Append( new Excel.Border() );                       /* 0 : 罫線なし */
            borders.Append(  CreateBorder( true, true, true, true ) );  /* 1 : 全罫線 */
            borders.Append( CreateBorder( true, false, true, true ) );  /* 2 : 下なし */
            borders.Append( CreateBorder( true, true, true, false ) );  /* 3 : 右なし */
            borders.Append( CreateBorder( true, false, true, false ) ); /* 4 : 下右なし */

            /* CellFormats */
            Excel.CellFormats cellFormats = new Excel.CellFormats();
            cellFormats.Append( new Excel.CellFormat() );
            for( uint i = 1; i <= 4; i++ )
            {
                cellFormats.Append( new Excel.CellFormat() { BorderId = i, ApplyBorder = true } );
            }

            Excel.Stylesheet stylesheet = new Excel.Stylesheet();
            stylesheet.Append( fonts );
            stylesheet.Append( fills );
            stylesheet.Append( borders );
            stylesheet.Append( cellFormats );

            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = stylesheet;
            stylesPart.Stylesheet.Save();
        }

        private static Excel.Border CreateBorder( bool a_TopBorder_flg, bool a_BottomBorder_flg, bool a_LeftBorder_flg, bool a_RightBorder_flg )
        {
            return new Excel.Border(
                a_LeftBorder_flg ? new Excel.LeftBorder() { Style = Excel.BorderStyleValues.Thin } : new Excel.LeftBorder(),
                a_RightBorder_flg ? new Excel.RightBorder(){ Style = Excel.BorderStyleValues.Thin } : new Excel.RightBorder(),
                a_TopBorder_flg ? new Excel.TopBorder() { Style = Excel.BorderStyleValues.Thin } : new Excel.TopBorder(),
                a_BottomBorder_flg ? new Excel.BottomBorder() { Style = Excel.BorderStyleValues.Thin } : new Excel.BottomBorder()
            );
        }

        private static void SetRow( Excel.SheetData sheetData, int rowIndex, List<CellData> values )
        {
            Excel.Row row = new Excel.Row();
            row.RowIndex = (uint)rowIndex;
            sheetData.Append( row );

            for( int i = 0; i < values.Count; i++ )
            {
                CellData data = values[i];

                Excel.Cell cell = new Excel.Cell();
                cell.CellReference = GetColumnName( i + 1 ) + rowIndex;
                cell.DataType = Excel.CellValues.String;
                cell.CellValue = new Excel.CellValue( data.Value ?? "" );

                /*
                 * StyleIndex決定
                 */
                uint style = 0;

                bool all =
                    data.BorderTop &&
                    data.BorderBottom &&
                    data.BorderLeft &&
                    data.BorderRight;

                bool noBottom =
                    data.BorderTop &&
                    !data.BorderBottom &&
                    data.BorderLeft &&
                    data.BorderRight;

                bool noRight =
                    data.BorderTop &&
                    data.BorderBottom &&
                    data.BorderLeft &&
                    !data.BorderRight;

                bool noBottomRight =
                    data.BorderTop &&
                    !data.BorderBottom &&
                    data.BorderLeft &&
                    !data.BorderRight;

                if( all )
                {
                    style = 1;
                }
                else if( noBottom )
                {
                    style = 2;
                }
                else if( noRight )
                {
                    style = 3;
                }
                else if( noBottomRight )
                {
                    style = 4;
                }
                else
                {
                    /* 処理なし */
                }

                cell.StyleIndex = style;
                row.Append( cell );
            }
        }

        private static string GetColumnName( int index )
        {
            string name = "";
            while( 0 < index )
            {
                index--;
                name = (char)( 'A' + ( index % 26 ) ) + name;
                index /= 26;
            }
            return name;
        }

        static string SafeSheetName( string name )
        {
            /* Excelシート名の禁止文字を半角スペースに置換 */
            char[] invalidid = { '\\', '/', '*', '[', ']', ':', '?', ',', '、', '／' };
            foreach( char c in invalidid )
            {
                name = name.Replace( c, ' ' );
            }

            /* 全角スペースを除去 */
            name = name.Replace( "　", "" );

            /* Excelシート名の長さ制限チェック */
            if( 31 < name.Length )
            {
                name = name.Substring( 0, 31 );
            }

            return string.IsNullOrWhiteSpace( name ) ? "Sheet" : name.Trim();
        }

        #endregion

        #region Word Helper

        class LevelDefinition
        {
            public Word.NumberFormatValues? Format;
            public string Text;
            public int Start;
        }

        class NumberingDefinition
        {
            public Dictionary<int, LevelDefinition> Levels = new Dictionary<int, LevelDefinition>();
        }

        /// <summary>
        /// 段落番号を生成するためのエンジンクラス
        /// Word の numbering.xml で定義された番号書式に基づき各レベルのカウンタを管理しながら番号文字列を生成する
        /// </summary>
        class NumberingEngine
        {
            /// <summary>
            /// 各レベルごとの現在のカウンタ値を保持する辞書
            /// キー  ： レベル番号（0,1,2,...)
            /// 値    ： そのレベルの現在の番号カウント
            /// </summary>
            private Dictionary<int, int> counters = new Dictionary<int, int>();

            /// <summary>
            /// 指定された番号定義およびレベルに基づいて現在の段落に対応する番号文字列を生成する
            /// </summary>
            /// <param name="def">番号定義情報。abstractNum 単位で定義されたレベル別の番号書式 (%1, %2 など) を保持する</param>
            /// <param name="level">対象となる段落のレベル (0 が最上位</param>
            /// <returns>生成された番号文字列 (例： "1.2", "1-3" など)</returns>
            public string Generate( NumberingDefinition def, int level )
            {
                string result = def.Levels[level].Text ?? "";

                if( result.Equals( "%1　" ) )
                {
                    return ConvertChapterNum( result.Trim() );
                }
                else if( result.Contains( "-" ) &&
                        1 < result.Length )
                {
                    return ConvertChapterNum( result.Trim() );
                }
                else
                {
                    return "";
                }
            }

            /// <summary>
            /// %1, %2, %3 を含む文字列を受け取り、階層番号に変換した結果を返す
            /// </summary>
            /// <param name="pattern"></param>
            /// <returns></returns>
            public string ConvertChapterNum( string pattern )
            {
                /* この行で使われている最大レベルを取得 */
                int maxLevel = GetMaxLevel( pattern );

                /* 該当レベルをインクリメント */
                if( !counters.ContainsKey( maxLevel ) )
                {
                    counters[maxLevel] = 0;
                }
                counters[maxLevel]++;

                /* 下位レベルをリセット */
                List<int> keys = new List<int>( counters.Keys );
                foreach( int k in keys )
                {
                    if( k > maxLevel )
                    {
                        counters[k] = 0;
                    }
                }

                /* %n を実際の番号に置き換え */
                string result = pattern;
                foreach( KeyValuePair<int, int> kv in counters )
                {
                    result = result.Replace( "%" + ( kv.Key + 1 ), kv.Value.ToString() );
                }

                /* 数字以外 (余分な % 等) を整理 */
                result = Regex.Replace( result, @"[^0-9\-]", "" );

                return result.Trim( '-' );
            }

            /// <summary>
            /// 文字列中に含まれる最大の %n を取得する
            /// </summary>
            /// <param name="pattern"></param>
            /// <returns></returns>
            private int GetMaxLevel( string pattern )
            {
                int max = 0;
                MatchCollection matches = Regex.Matches( pattern, @"%(\d+)" );

                foreach( Match m in matches )
                {
                    int level = int.Parse( m.Groups[1].Value )- 1;
                    if( max < level )
                    {
                        max = level;
                    }
                }
                return max;
            }

        }

        private static (int? numId, int? level ) GetNumberingInfo( Word.Paragraph pars, StyleDefinitionsPart stylePart )
        {
            Word.NumberingProperties numPr = pars.ParagraphProperties?.NumberingProperties;
            if( null != numPr )
            {
                return (
                    (int?)numPr.NumberingId?.Val?.Value,
                    (int?)numPr.NumberingLevelReference?.Val?.Value
                );
            }

            string styleId = pars.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if( null == styleId ) { return ( null, null ); }

            Word.Style style = stylePart?.Styles?.Elements<Word.Style>().FirstOrDefault( s => s.StyleId == styleId );
            Word.NumberingProperties styleNumPr = style?.StyleParagraphProperties?.NumberingProperties;
            return (
                (int?)styleNumPr?.NumberingId?.Val?.Value,
                (int?)styleNumPr?.NumberingLevelReference?.Val?.Value
            );
        }

        private static Dictionary<int, NumberingDefinition> LoadNumbring( WordprocessingDocument doc )
        {
            Dictionary<int, NumberingDefinition> result = new Dictionary<int, NumberingDefinition>();
            NumberingDefinitionsPart part = doc.MainDocumentPart.NumberingDefinitionsPart;
            if( null == part ) { return result; }
   
            Dictionary<int, NumberingDefinition> abstractMap = new Dictionary<int, NumberingDefinition>();

            foreach( Word.AbstractNum abs in part.Numbering.Elements<Word.AbstractNum>() )
            {
                NumberingDefinition def = new NumberingDefinition();
                foreach( Word.Level lvl in abs.Elements<Word.Level>() )
                {
                    def.Levels[(int)lvl.LevelIndex.Value] = new LevelDefinition
                    {
                        Text = null != lvl.LevelText ? lvl.LevelText.Val.Value : "",
                        Format = null != lvl.NumberingFormat ? lvl.NumberingFormat.Val.Value : (Word.NumberFormatValues?)null,
                        Start = null != lvl.StartNumberingValue ? lvl.StartNumberingValue.Val.Value : 1
                    };
                }
                abstractMap[ (int)abs.AbstractNumberId.Value ] = def;
            }

            foreach( Word.NumberingInstance num in part.Numbering.Elements<Word.NumberingInstance>() )
            {
                int numId = (int)num.NumberID.Value;
                int absId = (int)num.AbstractNumId.Val.Value;

                if( abstractMap.ContainsKey( absId ) )
                {
                    result[numId] = abstractMap[absId];
                }
            }

            return result;
        }

        private static string GetVisibleText( OpenXmlElement a_element )
        {
            StringBuilder sb = new StringBuilder();

            string currentField = null;

            foreach( OpenXmlElement element in a_element.Descendants() )
            {
                /* フィールド開始・終了 */
                Word.FieldChar fieldChar = element as Word.FieldChar;
                if( null != fieldChar )
                {
                    if( Word.FieldCharValues.End == fieldChar.FieldCharType )
                    {
                        currentField = null;
                    }

                    continue;
                }

                /* フィールドコード */
                Word.FieldCode fieldCode = element as Word.FieldCode;
                if( null != fieldCode )
                {
                    string txt = fieldCode.Text;

                    if( txt.Contains( "SEQ" ) )
                    {
                        currentField = "SEQ";
                    }
                    else if( txt.Contains( "STYLEREF" ) )
                    {
                        currentField = "STYLEREF";
                    }
                    else
                    {
                        /* 処理なし */
                    }

                    continue;
                }

                /* 表示テキスト */
                Word.Text text = element as Word.Text;
                if( null != text )
                {
                    string value = text.Text;

                    /* SEQ結果の前に "-" を補完 */
                    if( currentField == "SEQ" )
                    {
                        if( 0 < sb.Length &&
                            char.IsDigit( sb[sb.Length - 1] ) &&
                            0 < value.Length &&
                            char.IsDigit( value[0] ) )
                        {
                            sb.Append( "-" );
                        }
                    }

                    sb.Append( value );
                }
            }

            return sb.ToString();
        }

        #endregion

    }
}
