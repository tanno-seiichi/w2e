using System;

namespace w2e.file
{
    /// <summary>
    /// ファイルをコピーするクラス
    /// </summary>
    public static class FileCopy
    {
        /// <summary>
        /// ファイルの一時コピーを生成する
        /// </summary>
        /// <param name="a_filePath">コピーしたいファイル</param>
        /// <returns>一時コピーしたファイルのパス</returns>
        public static string CreateTempCopy( string a_filePath )
        {
            /* 元ファイルと同じフォルダではなく、OSのTempフォルダに作成する
             * （元フォルダが読み取り専用の場合でもコピーできるようにするため）
             */
            string dir = System.IO.Path.GetTempPath();
            string tempPath = System.IO.Path.Combine( dir, System.IO.Path.GetFileNameWithoutExtension( a_filePath ) + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + System.IO.Path.GetExtension( a_filePath ) );
            System.IO.File.Copy( a_filePath, tempPath, false );
            return tempPath;
        }


    }
}
