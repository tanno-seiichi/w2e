using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace w2e.word
{
    /// <summary>
    /// 段落番号を生成するためのエンジンクラス
    /// Word の numbering.xml で定義された番号書式に基づき各レベルのカウンタを管理しながら番号文字列を生成する
    /// </summary>
    public class NumberingEngine
    {
        /// <summary>
        /// 各レベルごとの現在のカウンタ値を保持する辞書
        /// キー  ： レベル番号（0,1,2,...)
        /// 値    ： そのレベルの現在の番号カウント
        /// </summary>
        private Dictionary<int, int> m_counters = new Dictionary<int, int>();

        /// <summary>
        /// 指定された番号定義およびレベルに基づいて現在の段落に対応する番号文字列を生成する
        /// </summary>
        /// <param name="a_def">Wordの章番号の番号付け(Numbering)定義情報。abstractNum 単位で定義されたレベル別の番号書式 (%1, %2 など) を保持する</param>
        /// <param name="a_level">対象となる段落のレベル (0 が最上位</param>
        /// <returns>生成された番号文字列 (例： "1.2", "1-3" など)</returns>
        public string Generate( NumberingDefinition a_def, int a_level )
        {
            /* Wordの"Numbering Level Text"を取得 */
            string result = a_def.Levels[a_level].text ?? "";

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
        /// Wordの章番号のレベル別の番号書式 (%1, %2 など)を受け取り、階層番号に変換した結果を返す
        /// </summary>
        /// <param name="a_pattern">Wordの章番号のレベル別の番号書式 (%1, %2 など)</param>
        /// <returns>階層番号に変換した文字列</returns>
        public string ConvertChapterNum( string a_pattern )
        {
            /* この行で使われている最大レベルを取得 */
            int maxLevel = GetMaxLevel( a_pattern );

            /* 該当レベルをインクリメント */
            if( !m_counters.ContainsKey( maxLevel ) )
            {
                m_counters[maxLevel] = 0;
            }
            m_counters[maxLevel]++;

            /* 下位レベルをリセット */
            List<int> keys = new List<int>( m_counters.Keys );
            foreach( int k in keys )
            {
                if( k > maxLevel )
                {
                    m_counters[k] = 0;
                }
            }

            /* %n を実際の番号に置き換え */
            string result = a_pattern;
            foreach( KeyValuePair<int, int> kv in m_counters )
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
        /// <param name="a_pattern">Wordの章番号のレベル別の番号書式 (%1, %2 など)</param>
        /// <returns>文字列中に含まれる最大の数値</returns>
        private int GetMaxLevel( string a_pattern )
        {
            int max = 0;
            MatchCollection matches = Regex.Matches( a_pattern, @"%(\d+)" );

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
}
