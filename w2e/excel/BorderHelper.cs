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
        /// <param name="a_rightAlign_flg">セル内の文字列を右揃えで表示するかどうか（箇条書きの記号「・」の表示などに使用する）</param>
        /// <param name="a_bold_flg">セル内の文字列を太字（ボールド）で表示するかどうか（章番号の見出し行の表示に使用する）</param>
        /// <param name="a_cache">スタイルキャッシュ（キー：罫線条件、値：StyleIndex）</param>
        public static void ApplyBorder( WorkbookPart a_wbPart, Cell a_cell, 
                                                 bool a_top_flg, bool a_bottom_flg, bool a_left_flg, bool a_right_flg, 
                                                 bool a_rightAlign_flg, bool a_bold_flg,
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

            /* 既存の元BorderIdと各辺のオンオフフラグ、右揃え設定、ボールド設定の組み合わせでキーを生成する */
            string key = baseBorderId.ToString() + "_" +
                 ( a_top_flg ? "1" : "0" ) +
                 ( a_bottom_flg ? "1" : "0" ) +
                 ( a_left_flg ? "1" : "0" ) +
                 ( a_right_flg ? "1" : "0" ) + "_" +
                 ( a_rightAlign_flg ? "1" : "0" ) +
                 ( a_bold_flg ? "1" : "0" );


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
                uint baseFontId = ( null != baseFormat.FontId ) ? baseFormat.FontId.Value : 0;

                /* フォントIDの決定（uintとUInt32Valueが混在する三項演算子はC# 7.3ではエラーになるためif文で分岐する） */
                if( a_bold_flg )
                {
                    newFormat.FontId = GetOrCreateHeadingFontId( styles, baseFontId );
                }
                else if( null != baseFormat.FontId )
                {
                    newFormat.FontId = baseFormat.FontId;
                }
                /* いずれにも該当しない場合はnewFormat.FontIdの既定値のままとする */

                newFormat.FillId = ( null != baseFormat.FillId ) ? baseFormat.FillId : newFormat.FillId;                                            /* 塗りつぶし */
                newFormat.NumberFormatId = ( null != baseFormat.NumberFormatId ) ? baseFormat.NumberFormatId : newFormat.NumberFormatId;            /* 数値書式 */
                newFormat.Alignment = ( null != baseFormat.Alignment ) ? ( Alignment)baseFormat.Alignment.CloneNode( true ) : newFormat.Alignment;  /* 配置 */
                newFormat.BorderId = newBorderId;                                                                                                   /* 罫線 */
                newFormat.ApplyBorder = true;                                                                                                       /* 罫線適用フラグをON */

                /* フォント適用フラグ（ボールド時のみON。bool/BooleanValueが混在する三項演算子はC# 7.3ではエラーになるためif文で分岐する） */
                if( a_bold_flg )
                {
                    newFormat.ApplyFont = true;
                }

                /* 枠線が設定されたセル、または右揃え指定があるセルは
                 * それぞれ必要な配置（枠線ありは折り返して全体を表示する／上詰め、右揃え指定は右詰め）を有効にする
                 */
                if( a_top_flg || a_bottom_flg || a_left_flg || a_right_flg || a_rightAlign_flg )
                {
                    Alignment newAlignment = new Alignment();

                    /* 枠線が設定されたセルは「折り返して全体を表示する」と「上詰め」を有効にする */
                    if( a_top_flg || a_bottom_flg || a_left_flg || a_right_flg )
                    {
                        newAlignment.WrapText = true;
                        newAlignment.Vertical = VerticalAlignmentValues.Top;
                    }

                    /* 右揃え指定があるセルは「右詰め」を有効にする（箇条書きの記号「・」の表示などに使用する） */
                    if( a_rightAlign_flg )
                    {
                        newAlignment.Horizontal = HorizontalAlignmentValues.Right;
                    }

                    newFormat.Alignment = newAlignment;
                    newFormat.ApplyAlignment = true;
                }

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


        /// <summary>
        /// フォントサイズの標準的な選択肢一覧（Excelの「フォントサイズ」欄のドロップダウンと同じ並び）。
        /// 「フォントサイズを大きくする」ボタンはこの一覧に沿ってサイズを変更するため、同じ動きを再現するために使用する。
        /// </summary>
        private static readonly double[] STANDARD_FONT_SIZES =
        {
            8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 96, 144
        };

        /// <summary>
        /// フォントサイズが不明な場合に基準とする既定値（Excelの標準フォントサイズ）
        /// </summary>
        private const double DEFAULT_FONT_SIZE = 11.0;

        /// <summary>
        /// 何段階サイズアップするか（見出し行の表示に使用する。Excelの「フォントサイズを大きくする」ボタンを2回押した時と同じ動き）
        /// </summary>
        private const int HEADING_FONT_SIZE_STEPS = 2;


        /// <summary>
        /// 見出し用フォント（ボールド＋サイズアップ）のIDを取得する。
        /// 既に同じ条件（ボールド・サイズ・フォント名）のフォントが登録済みの場合はそれを再利用し、無い場合は新規に作成する。
        /// </summary>
        /// <param name="a_styles">スタイルシート</param>
        /// <param name="a_baseFontId">元になるセルのフォントID（フォント名・サイズの基準として使用する）</param>
        /// <returns>見出し用フォントのFontId</returns>
        private static uint GetOrCreateHeadingFontId( Stylesheet a_styles, uint a_baseFontId )
        {
            List<Font> fonts = a_styles.Fonts.Elements<Font>().ToList();

            /* 基準となるフォント（サイズ・フォント名を引き継ぐ元）を取得する */
            Font baseFont = ( a_baseFontId < fonts.Count ) ? fonts[(int)a_baseFontId] : null;

            double baseSize = ( null != baseFont?.FontSize?.Val ) ? baseFont.FontSize.Val.Value : DEFAULT_FONT_SIZE;
            string baseFontName = baseFont?.FontName?.Val?.Value;

            /* Excelの「フォントサイズを大きくする」ボタンと同様に、標準サイズ一覧に沿って2段階分サイズを大きくする */
            double targetSize = IncreaseFontSize( baseSize, HEADING_FONT_SIZE_STEPS );

            /* 既に同じ条件（ボールド・対象サイズ・フォント名）のフォントが登録済であれば再利用する */
            for( int i = 0; i < fonts.Count; i++ )
            {
                Font f = fonts[i];
                bool isBold_flg = ( null != f.Bold );
                bool isSameSize_flg = ( null != f.FontSize ) && ( f.FontSize.Val.Value == targetSize );
                bool isSameName_flg = ( f.FontName?.Val?.Value == baseFontName );

                if( isBold_flg && isSameSize_flg && isSameName_flg )
                {
                    return (uint)i;
                }
            }

            /* 無い場合は新規に見出し用フォント（ボールド＋サイズアップ）を作成して追加する */
            Font newFont = new Font();
            newFont.Bold = new Bold();
            newFont.FontSize = new FontSize() { Val = targetSize };

            if( !string.IsNullOrEmpty( baseFontName ) )
            {
                /* 元のフォント名を引き継ぐ（未指定の場合はExcelの既定フォントに任せる） */
                newFont.FontName = new FontName() { Val = baseFontName };
            }

            a_styles.Fonts.AppendChild( newFont );
            uint newFontId = (uint)( a_styles.Fonts.Count() - 1 );
            a_styles.Fonts.Count = (uint)a_styles.Fonts.Count();

            return newFontId;
        }


        /// <summary>
        /// 標準フォントサイズ一覧に沿って、指定した段階数分サイズを大きくする
        /// </summary>
        /// <param name="a_currentSize">現在のフォントサイズ</param>
        /// <param name="a_steps">大きくする段階数</param>
        /// <returns>変更後のフォントサイズ</returns>
        private static double IncreaseFontSize( double a_currentSize, int a_steps )
        {
            /* 現在のサイズと同じか、それ以上で最も近いサイズのインデックスを探す */
            int index = System.Array.FindIndex( STANDARD_FONT_SIZES, s => s >= a_currentSize );
            if( 0 > index )
            {
                /* 一覧の最大値より大きい場合は最大値を基準にする */
                index = STANDARD_FONT_SIZES.Length - 1;
            }

            int targetIndex = System.Math.Min( index + a_steps, STANDARD_FONT_SIZES.Length - 1 );
            return STANDARD_FONT_SIZES[targetIndex];
        }
    }
}
