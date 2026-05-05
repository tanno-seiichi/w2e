using System.Collections.Generic;

namespace w2e.word
{
    /// <summary>
    /// Wordの章番号の番号付け（Numbering）定義を表すクラス
    /// </summary>
    public class NumberingDefinition
    {
        /* OpenXMLにおける「AbstractNum（抽象番号定義）」に対応する情報を保持します
         * 各レベル（Level 0, 1, 2...）ごとの番号書式や開始値などを <see cref="LevelDefinition"/> として管理します。
         * 
         * 主に以下の用途で使用されます：
         * ・NumberingDefinitionsPart から読み込んだ番号情報の保持
         * ・段落（Paragraph）の番号表示形式の解決
         * 
         * 注意：
         * ・キー（int）はレベル番号（LevelIndex）を表します
         * ・存在しないレベルへのアクセスには注意が必要です
         */

        /// <summary>
        /// 各レベルごとの番号定義を保持する辞書
        /// </summary>
        /// <remarks>
        /// キー：レベル番号（0始まり）
        /// 値：そのレベルに対応する番号定義
        /// 
        /// 例：
        /// ・0 → 第1階層（章など）
        /// ・1 → 第2階層（節など）
        /// ・2 → 第3階層（項など）
        public Dictionary<int, LevelDefinition> Levels { get; } = new Dictionary<int, LevelDefinition>();

    }
}
