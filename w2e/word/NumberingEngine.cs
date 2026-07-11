using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
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
            string result = "";
            if( a_def.Levels.ContainsKey( a_level ) )
            {
                result = a_def.Levels[a_level].text ?? "";
            }

            if( result.Equals( "%1　" ) ||
                result.Equals( "%1" ) )
            {
                return ConvertChapterNum( result.Trim() );
            }
            else if( ( result.Contains( "-" ) 
                    || result.Contains( ".") ) &&
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
            result = Regex.Replace( result, @"[^0-9\-.]", "" );

            return result.Trim( '-' ).Trim( '.' );
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


        /// <summary>
        /// 箇条書き（見出し以外の番号付き段落）の記号文字列を生成する。
        /// 章番号(Generate)とは異なり、"1." "①" "a)" "・" など、Wordのフォーマット定義通りの記号を生成する。
        /// </summary>
        /// <param name="a_def">Wordの番号付け(Numbering)定義情報</param>
        /// <param name="a_level">対象となる段落のレベル (0 が最上位)</param>
        /// <param name="a_counters">
        /// このリスト（numId）専用の、レベルごとの現在のカウンタ値を保持する辞書。
        /// 呼び出し側でnumIdごとに個別のインスタンスを用意し、他の番号付けと状態を共有しないようにすること
        /// </param>
        /// <returns>生成された記号文字列 (例： "1.", "①", "a)", "・" など)</returns>
        public string GenerateListMarker( NumberingDefinition a_def, int a_level, Dictionary<int, int> a_counters )
        {
            if( !a_def.Levels.ContainsKey( a_level ) )
            {
                return "";
            }

            LevelDefinition levelDef = a_def.Levels[a_level];

            /* このレベルのカウンタを更新する（初回は開始値、以降はインクリメント） */
            if( !a_counters.ContainsKey( a_level ) )
            {
                a_counters[a_level] = levelDef.start;
            }
            else
            {
                a_counters[a_level]++;
            }

            /* 下位レベルのカウンタは削除しておく（次に出現した時に開始値から始まるようにするため） */
            List<int> deeperKeys = a_counters.Keys.Where( k => k > a_level ).ToList();
            foreach( int k in deeperKeys )
            {
                a_counters.Remove( k );
            }

            /* テンプレート中の %1, %2, ... を、該当レベルの現在のカウンタ値（書式変換後）に置き換える
             * テンプレートに %n が含まれない場合（Bulletなど記号そのものが設定されている場合）はそのまま返る
             */
            string result = levelDef.text ?? "";

            for( int l = 0; l <= a_level; l++ )
            {
                string placeholder = "%" + ( l + 1 );
                if( !result.Contains( placeholder ) ) { continue; }

                int counterValue = a_counters.ContainsKey( l ) ? a_counters[l] :
                                    ( a_def.Levels.ContainsKey( l ) ? a_def.Levels[l].start : 1 );
                NumberFormatValues? levelFormat = a_def.Levels.ContainsKey( l ) ? a_def.Levels[l].format : null;

                result = result.Replace( placeholder, FormatNumberValue( counterValue, levelFormat ) );
            }

            return result;
        }


        /// <summary>
        /// 数値を指定された番号フォーマットに応じた文字列に変換する
        /// </summary>
        /// <param name="a_value">変換対象の数値</param>
        /// <param name="a_format">Wordの番号フォーマット種別（null の場合は10進数として扱う）</param>
        /// <returns>フォーマット変換後の文字列</returns>
        private static string FormatNumberValue( int a_value, NumberFormatValues? a_format )
        {
            /* OpenXml 3.x以降、NumberFormatValuesは列挙型(enum)ではなく構造体になっており、
             * switch文のcaseラベルに指定できない（コンパイルエラーになる）ため、if文で比較する
             */
            if( NumberFormatValues.DecimalEnclosedCircle == a_format )
            {
                return ToCircledNumber( a_value );
            }
            else if( NumberFormatValues.LowerLetter == a_format )
            {
                return ToAlpha( a_value, false );
            }
            else if( NumberFormatValues.UpperLetter == a_format )
            {
                return ToAlpha( a_value, true );
            }
            else if( NumberFormatValues.LowerRoman == a_format )
            {
                return ToRoman( a_value ).ToLowerInvariant();
            }
            else if( NumberFormatValues.UpperRoman == a_format )
            {
                return ToRoman( a_value );
            }
            else
            {
                /* Decimal およびその他未対応の書式は10進数として扱う */
                return a_value.ToString();
            }
        }


        /// <summary>
        /// 数値を丸数字（①〜⑳）に変換する。範囲外の場合は "(n)" の形式で代用する
        /// </summary>
        /// <param name="a_value">変換対象の数値</param>
        /// <returns>丸数字の文字列</returns>
        private static string ToCircledNumber( int a_value )
        {
            /* Unicodeの丸数字は ① (U+2460) 〜 ⑳ (U+2473) の20個のみ定義されている */
            if( 1 <= a_value && a_value <= 20 )
            {
                return char.ConvertFromUtf32( 0x2460 + ( a_value - 1 ) );
            }

            return "(" + a_value + ")";
        }


        /// <summary>
        /// 数値をアルファベット（a, b, ... z, aa, ab, ...）に変換する
        /// </summary>
        /// <param name="a_value">変換対象の数値（1始まり）</param>
        /// <param name="a_upper">大文字にするかどうか</param>
        /// <returns>アルファベットの文字列</returns>
        private static string ToAlpha( int a_value, bool a_upper )
        {
            string result = "";
            int n = a_value;

            while( 0 < n )
            {
                n--;
                result = (char)( ( a_upper ? 'A' : 'a' ) + ( n % 26 ) ) + result;
                n /= 26;
            }

            return result;
        }


        /// <summary>
        /// 数値をローマ数字（大文字）に変換する
        /// </summary>
        /// <param name="a_value">変換対象の数値</param>
        /// <returns>ローマ数字の文字列</returns>
        private static string ToRoman( int a_value )
        {
            if( a_value <= 0 )
            {
                return a_value.ToString();
            }

            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };

            int remaining = a_value;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            for( int i = 0; i < values.Length; i++ )
            {
                while( remaining >= values[i] )
                {
                    sb.Append( symbols[i] );
                    remaining -= values[i];
                }
            }

            return sb.ToString();
        }


    }
}
