using System.IO;
using System.Linq;
using System.Text;

namespace w2e.markdown
{
    /// <summary>
    /// MarkDownファイルを書き出すクラス
    /// </summary>
    public class MarkDownWriter
    {

        /// <summary>
        /// 文字列構築オブジェクト
        /// </summary>
        private StringBuilder m_sb = new StringBuilder();

        /// <summary>
        /// 現在のファイル
        /// </summary>
        private string m_currentFile;

        /// <summary>
        /// ファイル名の禁止文字を除去して返す
        /// </summary>
        /// <param name="a_name">禁止文字を除去する前のファイル名</param>
        /// <returns>禁止文字を除去したファイル名</returns>
        public static string SafeFileName( string a_name )
        {
            /* ファイル名の禁止文字を半角スペースに置換 */
            char[] invalidid = { '\\', '/', '*', '<', '>', ':', '?', ',', '、', '／' };
            foreach( char c in invalidid )
            {
                a_name = a_name.Replace( c, ' ' );
            }

            /* 全角スペースを除去 */
            a_name = a_name.Replace( "　", "" );

            return string.IsNullOrWhiteSpace( a_name ) ? "File" : a_name.Trim();
        }

        /// <summary>
        /// 現在のファイルを書き出してから新しいファイル用に情報を初期化します
        /// </summary>
        /// <param name="a_filePath">新しいファイルパス</param>
        public void NewFile( string a_filePath )
        {
            Save(); // 前のファイル保存

            m_currentFile = a_filePath;
            m_sb.Clear();
        }


        /// <summary>
        /// 行を追加します
        /// </summary>
        /// <param name="a_text">行に記述する文字列</param>
        public void AddLine( string a_text )
        {
            m_sb.AppendLine( a_text );
        }


        /// <summary>
        /// テーブル行を追加します
        /// </summary>
        /// <param name="a_cols">行に追加する列の配列</param>
        public void AddTableRow( params string[] a_cols )
        {
            m_sb.AppendLine( "| " + string.Join( " | ", a_cols ) + " |" );
        }


        /// <summary>
        /// テーブルのヘッダ行とデータ行のセパレータを追加します
        /// </summary>
        /// <param name="a_colCount">テーブルの列数</param>
        public void AddTableSeparator( int a_colCount )
        {
            m_sb.AppendLine( "| " + string.Join( " | ", Enumerable.Repeat( "---", a_colCount ) ) + " |" );
        }


        /// <summary>
        /// 現在のファイルを書き出します
        /// </summary>
        public void Save()
        {
            if( !string.IsNullOrEmpty( m_currentFile ) )
            {
                File.WriteAllText( m_currentFile, m_sb.ToString(), Encoding.UTF8 );
            }
        }


    }
}
