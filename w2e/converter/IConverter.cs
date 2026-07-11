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
        /// <param name="a_outputImage_flg">画像を使用するか否か</param>
        /// <param name="a_outputListNumber_flg">箇条書きに番号（Wordの実際の記号）を使用するか否か。falseの場合は固定の記号（Excelは「・」、MarkDownは「-」）を使用する</param>
        /// <param name="a_token">処理中断通知</param>
        void Convert( string a_inputPath, string a_outputPath, bool a_outputImage_flg, bool a_outputListNumber_flg, CancellationToken a_token );

    }
}
