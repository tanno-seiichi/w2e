using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace w2e.excel
{
    public static class ExcelHelper
    {
        public static WorksheetPart CreateWorksheet( WorkbookPart a_wbPart, Sheets a_sheets, string a_sheetName, uint a_sheetId, out SheetData a_sheetData )
        {
            WorksheetPart wsPart = a_wbPart.AddNewPart<WorksheetPart>();

            a_sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet( a_sheetData );

            Sheet sheet = new Sheet();
            sheet.Id = a_wbPart.GetIdOfPart( wsPart );
            sheet.SheetId = a_sheetId;
            sheet.Name = a_sheetName;

            a_sheets.Append( sheet );
            return wsPart;
        }

        /// <summary>
        /// WorkbookStylesPart と Stylesheet を安全に初期化する。
        /// Excelが要求する最小構造（Fonts/Fills/Borders/CellFormats）を必ず1件以上持たせる。
        /// </summary>
        /// <param name="a_wbPart">WorkbookPart</param>
        public static void InitializeStylesheet( WorkbookPart a_wbPart )
        {
            /* WorkbookStylesPart取得（無ければ作成） */
            WorkbookStylesPart stylesPart = a_wbPart.WorkbookStylesPart;
            if( null == stylesPart )
            {
                stylesPart = a_wbPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
            }

            /* Stylesheet取得（無ければ生成） */
            Stylesheet styles = stylesPart.Stylesheet;
            if( null == styles )
            {
                styles = new Stylesheet();
                stylesPart.Stylesheet = styles;
            }

            /* Fonts初期化（最低1件必要） */
            if( null == styles.Fonts )
            {
                styles.Fonts = new Fonts( new Font() );
            }
            else if( styles.Fonts.Count() == 0 )
            {
                styles.Fonts.AppendChild( new Font() );
            }

            /* Fills初期化（最低1件必要） */
            if( null == styles.Fills )
            {
                styles.Fills = new Fills( new Fill() );
            }
            else if( styles.Fills.Count() == 0 )
            {
                styles.Fills.AppendChild( new Fill() );
            }

            /* Borders初期化（最低1件必要） */
            if( null == styles.Borders )
            {
                styles.Borders = new Borders( new Border() );
            }
            else if( styles.Borders.Count() == 0 )
            {
                styles.Borders.AppendChild( new Border() );
            }

            /* CellFormats初期化（最低1件必要） */
            if( null == styles.CellFormats )
            {
                styles.CellFormats = new CellFormats( new CellFormat() );
            }
            else if( styles.CellFormats.Count() == 0 )
            {
                styles.CellFormats.AppendChild( new CellFormat() );
            }

            /* Count属性を整合させる */
            styles.Fonts.Count = (uint)styles.Fonts.Count();
            styles.Fills.Count = (uint)styles.Fills.Count();
            styles.Borders.Count = (uint)styles.Borders.Count();
            styles.CellFormats.Count = (uint)styles.CellFormats.Count();
        }

#if BORDER_1
        public static void CreateStylesheet( WorkbookPart a_workbookPart )
        {

            Fonts fonts = new Fonts( new Font() );
            Fills fills = new Fills( new Fill( new PatternFill() ) );

            Borders borders = new Borders();
            borders.Append( new Border() );                       /* 0 : 罫線なし */
            borders.Append( CreateBorder( true, true, true, true ) );  /* 1 : 全罫線 */
            borders.Append( CreateBorder( true, false, true, true ) );  /* 2 : 下なし */
            borders.Append( CreateBorder( true, true, true, false ) );  /* 3 : 右なし */
            borders.Append( CreateBorder( true, false, true, false ) ); /* 4 : 下右なし */

            /* CellFormats */
            CellFormats cellFormats = new CellFormats();
            cellFormats.Append( new CellFormat() );
            for( uint i = 1; i <= 4; i++ )
            {
                cellFormats.Append( new CellFormat() { BorderId = i, ApplyBorder = true } );
            }

            Stylesheet stylesheet = new Stylesheet();
            stylesheet.Append( fonts );
            stylesheet.Append( fills );
            stylesheet.Append( borders );
            stylesheet.Append( cellFormats );

            WorkbookStylesPart stylesPart = a_workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = stylesheet;
            stylesPart.Stylesheet.Save();
        }

        private static Border CreateBorder( bool a_TopBorder_flg, bool a_BottomBorder_flg, bool a_LeftBorder_flg, bool a_RightBorder_flg )
        {
            return new Border(
                a_LeftBorder_flg ? new LeftBorder() { Style = BorderStyleValues.Thin } : new LeftBorder(),
                a_RightBorder_flg ? new RightBorder() { Style = BorderStyleValues.Thin } : new RightBorder(),
                a_TopBorder_flg ? new TopBorder() { Style = BorderStyleValues.Thin } : new TopBorder(),
                a_BottomBorder_flg ? new BottomBorder() { Style = BorderStyleValues.Thin } : new BottomBorder()
            );
        }

        public static void SetRow( SheetData a_sheetData, int a_rowIndex, List<CellData> a_values )
        {
            Row row = new Row();
            row.RowIndex = (uint)a_rowIndex;
            a_sheetData.Append( row );

            for( int i = 0; i < a_values.Count; i++ )
            {
                CellData data = a_values[i];

                Cell cell = new Cell();
                cell.CellReference = GetColumnName( i + 1 ) + a_rowIndex;
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue( data.Value ?? "" );

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
#else
        public static void SetRow( WorkbookPart a_wbPart, SheetData a_sheetData, int a_rowIndex, List<CellData> a_values, Dictionary<string, uint> a_cache )
        {
            Row row = new Row();
            row.RowIndex = (uint)a_rowIndex;
            a_sheetData.Append( row );

            for( int i = 0; i < a_values.Count; i++ )
            {
                CellData data = a_values[i];

                Cell cell = new Cell();
                cell.CellReference = GetColumnName( i + 1 ) + a_rowIndex;
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue( data.Value ?? "" );

                BorderHelper.ApplyBorderWithCache( a_wbPart, cell, data.BorderTop, data.BorderBottom, data.BorderLeft, data.BorderRight, a_cache );

                row.Append( cell );
            }
        }
#endif


        private static string GetColumnName( int a_index )
        {
            string name = "";
            while( 0 < a_index )
            {
                a_index--;
                name = (char)( 'A' + ( a_index % 26 ) ) + name;
                a_index /= 26;
            }
            return name;
        }

        public static string SafeSheetName( string a_name )
        {
            /* Excelシート名の禁止文字を半角スペースに置換 */
            char[] invalidid = { '\\', '/', '*', '[', ']', ':', '?', ',', '、', '／' };
            foreach( char c in invalidid )
            {
                a_name = a_name.Replace( c, ' ' );
            }

            /* 全角スペースを除去 */
            a_name = a_name.Replace( "　", "" );

            /* Excelシート名の長さ制限チェック */
            if( 31 < a_name.Length )
            {
                a_name = a_name.Substring( 0, 31 );
            }

            return string.IsNullOrWhiteSpace( a_name ) ? "Sheet" : a_name.Trim();
        }

    }
}
