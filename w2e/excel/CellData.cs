namespace w2e.excel
{
    /// <summary>
    /// Excelのセルの内容を保持するクラス
    /// </summary>
    public class CellData
    {
        /// <summary>
        /// セルに表示する文字列
        /// </summary>
        public string text { get; set; } = "";
        
        /// <summary>
        /// 枠線（上）
        /// </summary>
        public bool topBorder { get; set; } = false;
        
        /// <summary>
        /// 枠線（下）
        /// </summary>
        public bool bottomBorder { get; set; } = false;
        
        /// <summary>
        /// 枠線（左）
        /// </summary>
        public bool leftBorder { get; set; } = false;

        /// <summary>
        /// 枠線（右）
        /// </summary>
        public bool rightBorder { get; set; } = false;

        /// <summary>
        /// セル内の文字列を右揃えで表示するかどうか（箇条書きの記号「・」の表示などに使用する）
        /// </summary>
        public bool rightAlign { get; set; } = false;

        /// <summary>
        /// セル内の文字列を太字（ボールド）で表示するかどうか（章番号の見出し行の表示に使用する）
        /// </summary>
        public bool bold { get; set; } = false;

    }
}
