using System;
using System.Collections.Generic;

namespace w2e.word
{
    /// <summary>
    /// 画像ファイル名を生成するクラス
    /// </summary>
    public class ImageFileNameGenerator
    {
        /* 章ごとの画像インデックス */
        private readonly Dictionary<string, int> m_imageIndexMap = new Dictionary<string, int>();


        /// <summary>
        /// 画像ファイル名を生成する。
        /// </summary>
        /// <param name="a_headingNumber">章番号</param>
        /// <param name="a_contentType">コンテンツタイプ</param>
        /// <returns>画像ファイル名</returns>
        public string CreateFileName( string a_headingNumber, string a_contentType )
        {
            /* 引数チェック */
            if( null == a_headingNumber )
            {
                throw new ArgumentNullException( nameof( a_headingNumber ) );
            }

            if( null == a_contentType )
            {
                throw new ArgumentNullException( nameof( a_contentType ) );
            }

            /* 章番号を正規化する */
            string normalizedHeadingNumber = NormalizeHeadingNumber( a_headingNumber );

            /* 章ごとのインデックス取得 */
            if( false == m_imageIndexMap.ContainsKey( normalizedHeadingNumber ) )
            {
                m_imageIndexMap[normalizedHeadingNumber] = 0;
            }

            int index = m_imageIndexMap[ normalizedHeadingNumber ];

            /* インクリメント */
            m_imageIndexMap[normalizedHeadingNumber] = index + 1;

            /* 拡張子取得 */
            string extension = GetExtension( a_contentType );

            /* ファイル名生成 */
            string fileName = "image_" + normalizedHeadingNumber + "_" + index.ToString( "0000" ) + extension;

            return fileName;
        }


        /// <summary>
        /// コンテンツタイプから拡張子を取得する。
        /// </summary>
        /// <param name="a_contentType">コンテンツタイプ</param>
        /// <returns>拡張子</returns>
        private string GetExtension( string a_contentType )
        {
            /* コンテンツタイプに応じて拡張子を返す */
            if( "image/png" == a_contentType )
            {
                return ".png";
            }

            if( "image/jpeg" == a_contentType )
            {
                return ".jpg";
            }

            if( "image/gif" == a_contentType )
            {
                return ".gif";
            }

            if( "image/bmp" == a_contentType )
            {
                return ".bmp";
            }

            if( "image/tiff" == a_contentType )
            {
                return ".tiff";
            }

            return ".bin";
        }


        /// <summary>
        /// 章番号をファイル名用に正規化する。
        /// </summary>
        /// <param name="a_headingNumber">章番号</param>
        /// <returns>正規化された章番号</returns>
        private string NormalizeHeadingNumber( string a_headingNumber )
        {
            /* 文字列を安全な形式に変換する */
            string result = a_headingNumber;

            result = result.Replace( ".", "-" );
            result = result.Replace( "/", "-" );
            result = result.Replace( " ", "" );

            return result;
        }
 
    
    }
}