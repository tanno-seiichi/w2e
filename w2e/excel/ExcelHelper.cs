using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace w2e.excel
{
    /// <summary>
    /// Excelを操作するクラス
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// Excelのワークシートを生成する
        /// </summary>
        /// <param name="a_wbPart">ワークシートの追加先となるブック</param>
        /// <param name="a_sheets">ブック内のシートコレクション</param>
        /// <param name="a_sheetName">作成するワークシートの名前</param>
        /// <param name="a_sheetId">ワークシートに割り当てるID（ブック内で一意である必要がある）</param>
        /// <param name="a_sheetData">ワークシートの内容</param>
        /// <returns> Excelワークシート</returns>
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
        /// Excelシート名の禁止文字を除去して返す
        /// </summary>
        /// <param name="a_name">禁止文字を除去する前のシート名</param>
        /// <returns>禁止文字を除去したシート名</returns>
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


        /// <summary>
        /// WorkbookStylesPart と Stylesheet を安全に初期化する。
        /// Excelが要求する最小構造（Fonts/Fills/Borders/CellFormats）を必ず1件以上持たせる。
        /// </summary>
        /// <param name="a_wbPart">WorkbookPart（スタイル情報のルート）</param>
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


        /// <summary>
        /// Excelの列名を取得する
        /// </summary>
        /// <param name="a_index">列インデックス（例：1 = A、26 = Z、27 = AA）</param>
        /// <returns>Excel形式の列名文字列</returns>
        private static string GetColumnName( int a_index )
        {
            /* 前提：Excelは 26進数、ただし0がない（A=1） */

            string name = "";   // 生成する列名

            /* インデックスを26進数（A-Z）として扱い、右側の桁から順に算出する */
            while( 0 < a_index )
            {
                /* 0始まりに補正（A=0として扱うため） */
                a_index--;

                /* 現在の桁の文字を算出して先頭に追加 */
                name = (char)( 'A' + ( a_index % 26 ) ) + name;

                /* 次の桁へ */
                a_index /= 26;
            }
            return name;
        }


        /// <summary>
        /// Excelの行を追加する
        /// </summary>
        /// <param name="a_wbPart">WorkbookPart（スタイル情報のルート）</param>
        /// <param name="a_sheetData">Excelワークシート</param>
        /// <param name="a_rowIndex">行番号</param>
        /// <param name="a_values">行に設定するセルデータの一覧</param>
        /// <param name="a_cache">Excelのスタイルシートに登録済のスタイルを再利用するためのキャッシュ</param>
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

                /* セルに値を設定 */
                if( string.IsNullOrEmpty( data.text ) )
                {
                    /* セルがブランクの場合 */
                    cell.DataType = null;
                    cell.CellValue = null;
                }
                else
                {
                    /* セルがブランクでない場合 */
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue( data.text ?? "" );
                }

                BorderHelper.ApplyBorder( a_wbPart, cell, data.topBorder, data.bottomBorder, data.leftBorder, data.rightBorder, a_cache );

                row.Append( cell );
            }
        }


    }
}
