using System.Threading;
using w2e.delegates;

namespace w2e.converter
{
    /// <summary>
    /// 変換処理を実行するクラスのインターフェース
    /// </summary>
    interface IConverter
    {
        /// <summary>
        /// 進捗情報が更新された時の処理
        /// </summary>
        Delegates.UpdateProgressDelegate onProgressUpdate { get; set; }

        /// <summary>
        /// ログが出力された時の処理
        /// </summary>
        Delegates.UpdateLogDelegate onLogUpdate { get; set; }

        /// <summary>
        /// 変換処理を実行する
        /// </summary>
        /// <param name="a_inputPath">入力パス</param>
        /// <param name="a_outputPath">出力パス</param>
        /// <param name="a_token">処理中断通知</param>
        void Convert( string a_inputPath, string a_outputPath, CancellationToken a_token );

    }
}
