using System;

namespace w2e.file
{
    public static class FileCopy
    {
        public static string CreateTempCopy( string a_wordPath )
        {
            string dir = System.IO.Path.GetDirectoryName( a_wordPath );
            string tempPath = System.IO.Path.Combine( dir, System.IO.Path.GetFileNameWithoutExtension( a_wordPath ) + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + System.IO.Path.GetExtension( a_wordPath ) );
            System.IO.File.Copy( a_wordPath, tempPath, false );
            return tempPath;
        }
    }
}
