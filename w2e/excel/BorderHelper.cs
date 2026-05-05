using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace w2e.excel
{
    /// <summary>
    /// Excelのセルに枠線を設定するクラス
    /// </summary>
    public static class BorderHelper
    {
        /// <summary>
        /// 既存スタイル（フォント・塗りつぶし・表示形式など）を維持したまま、罫線のみを変更して適用する。
        /// 同一条件のスタイルはキャッシュして再利用し、ファイル肥大化と処理時間増加を防止する。
        /// </summary>
        /// <param name="a_wbPart">WorkbookPart（スタイル情報のルート）</param>
        /// <param name="a_cell">対象セル</param>
        /// <param name="a_top_flg">上罫線を引くかどうか</param>
        /// <param name="a_bottom_flg">下罫線を引くかどうか</param>
        /// <param name="a_left_flg">左罫線を引くかどうか</param>
        /// <param name="a_right_flg">右罫線を引くかどうか</param>
        /// <param name="a_cache">スタイルキャッシュ（キー：罫線条件、値：StyleIndex）</param>
        public static void ApplyBorderWithCache( WorkbookPart a_wbPart, Cell a_cell, 
                                                 bool a_top_flg, bool a_bottom_flg, bool a_left_flg, bool a_right_flg, 
                                                 Dictionary<string, uint> a_cache )
        {
            /* スタイルシートが初期化されていない場合は何もしないで返す */
            if( null == a_wbPart?.WorkbookStylesPart ) { return; }

            /* スタイルシートを取得 */
            Stylesheet styles = a_wbPart.WorkbookStylesPart.Stylesheet;

            /* セルに設定済のスタイル "CellFormat" を取得する */
            /* スタイルが未設定、またはスタイル・インデックスが不正の場合は空のCellFormatを使用する */
            CellFormat baseFormat = null;
            uint baseFormatIndex = 0;
            if( ( null != a_cell.StyleIndex ) &&
                ( null != styles.CellFormats ) &&
                ( styles.CellFormats.Count() > a_cell.StyleIndex.Value ) )
            {
                baseFormatIndex = a_cell.StyleIndex.Value;
                baseFormat = (CellFormat)styles.CellFormats.ElementAt( (int)baseFormatIndex );
            }
            else
            {
                baseFormat = new CellFormat();
            }

            /* 既存のBorderIdを取得（未設定なら0） */
            uint baseBorderId = ( null != baseFormat.BorderId ) ? baseFormat.BorderId.Value : 0;

            /* ================== キャッシュキーを生成 ================== */

            /* 既存の元BorderIdと各辺のオンオフフラグの組み合わせでキーを生成する */
            string key = baseBorderId.ToString() + "_" +
                 ( a_top_flg ? "1" : "0" ) +
                 ( a_bottom_flg ? "1" : "0" ) +
                 ( a_left_flg ? "1" : "0" ) +
                 ( a_right_flg ? "1" : "0" );


            /* スタイルのキーがキャッシュに存在しない場合は生成して登録する */
            if( !a_cache.ContainsKey( key ) )
            {
                /* ================== 新しいBorder（罫線定義）を作成 ================== */

                /* 4辺の罫線をまとめてBorderオブジェクトを生成 */
                LeftBorder leftBorder       = a_left_flg    ? new LeftBorder() { Style = BorderStyleValues.Thin }   : new LeftBorder();     /* 左罫線 */
                RightBorder rightBorder     = a_right_flg   ? new RightBorder() { Style = BorderStyleValues.Thin }  : new RightBorder();    /* 右罫線 */
                TopBorder topBorder         = a_top_flg     ? new TopBorder() { Style = BorderStyleValues.Thin }    : new TopBorder();      /* 上罫線 */
                BottomBorder bottomBorder   = a_bottom_flg  ? new BottomBorder() { Style = BorderStyleValues.Thin } : new BottomBorder();   /* 下罫線 */
                Border newBorder = new Border(leftBorder, rightBorder, topBorder, bottomBorder, new DiagonalBorder());

                /* Bordersコレクションに追加（Excel内部の罫線定義として登録） */
                /* 追加したBorderのインデックスを取得 */
                styles.Borders.AppendChild( newBorder );
                uint newBorderId = (uint)(styles.Borders.Count() - 1);


                /* =============== CellFormat（セルの見た目定義）を作成 =============== */

                CellFormat newFormat = new CellFormat();
                newFormat.FontId = ( null != baseFormat.FontId ) ? baseFormat.FontId : newFormat.FontId;                                            /* フォント */
                newFormat.FillId = ( null != baseFormat.FillId ) ? baseFormat.FillId : newFormat.FillId;                                            /* 塗りつぶし */
                newFormat.NumberFormatId = ( null != baseFormat.NumberFormatId ) ? baseFormat.NumberFormatId : newFormat.NumberFormatId;            /* 数値書式 */
                newFormat.Alignment = ( null != baseFormat.Alignment ) ? ( Alignment)baseFormat.Alignment.CloneNode( true ) : newFormat.Alignment;  /* 配置 */
                newFormat.BorderId = newBorderId;                                                                                                   /* 罫線 */
                newFormat.ApplyBorder = true;                                                                                                       /* 罫線適用フラグをON */

                /* CellFormatsコレクションに追加 */
                /* 新しいスタイルのインデックスを取得 */
                styles.CellFormats.AppendChild( newFormat );
                uint newFormatId = (uint)(styles.CellFormats.Count() - 1);


                /* =============== キャッシュに登録 =============== */

                a_cache[key] = newFormatId;
            }

            /* セルにスタイルを適用 */
            a_cell.StyleIndex = a_cache[key];
        }


    }
}
