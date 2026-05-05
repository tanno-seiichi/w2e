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

    }
}
