namespace w2e.delegates
{
    /// <summary>
    /// Delegateを定義するクラス
    /// </summary>
    public static class Delegates
    {
        /// <summary>
        /// 進捗情報の処理を委譲するDelegate
        /// </summary>
        /// <param name="a_value"></param>
        public delegate void UpdateProgressDelegate( int a_value );

        /// <summary>
        /// ログ出力処理をを委譲するDelegate
        /// </summary>
        /// <param name="a_value"></param>
        public delegate void UpdateLogDelegate( string a_value );

    }
}
