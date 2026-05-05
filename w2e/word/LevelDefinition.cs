using DocumentFormat.OpenXml.Wordprocessing;

namespace w2e.word
{
    /// <summary>
    /// Wordの番号付けにおける「レベル単位」の定義を表すクラス
    /// </summary>
    public class LevelDefinition
    {
        /* OpenXMLの Level 要素に対応し、特定の階層における番号表示形式を保持します
         * 
         * 例：
         * ・第1階層：1, 2, 3...
         * ・第2階層：1.1, 1.2, 1.3...
         * ・第3階層：(a), (b), (c)...
         * 
         * 本クラスは以下の情報を管理します：
         * ・番号の表示テンプレート（Text）
         * ・番号フォーマット（Format）
         * ・開始番号（Start）
         * 
         * これらの情報をもとに、実際の表示文字列を生成する際に使用されます
         */

        /// <summary>
        /// 番号の表示テンプレート。
        /// </summary>
        /// <remarks>
        /// OpenXMLの LevelText に対応します
        ///
        /// 例：
        /// ・"%1."      → "1.", "2.", "3."
        /// ・"%1.%2."   → "1.1.", "1.2."
        /// 
        /// %n はレベル番号のプレースホルダを表します
        /// </remarks>
        public string text { get; set; }

        /// <summary>
        /// 番号のフォーマット種別
        /// </summary>
        /// <remarks>
        /// OpenXMLの NumberingFormat に対応します
        ///
        /// 例：
        /// ・Decimal      → 1, 2, 3
        /// ・LowerLetter  → a, b, c
        /// ・UpperRoman   → I, II, III
        /// 
        /// null の場合、フォーマット未指定を意味します
        /// </remarks>
        public NumberFormatValues? format { get; set; }

        /// <summary>
        /// 番号の開始値。
        /// </summary>
        /// <remarks>
        /// OpenXMLの StartNumberingValue に対応します
        /// 
        /// 通常は 1 が指定されますが、途中から開始する場合などに変更されることがあります
        /// </remarks>
        public int start { get; set; }

    }
}
