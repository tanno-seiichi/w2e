using Microsoft.Win32;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using w2e.converter;

namespace w2e
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    /// <remarks>
    /// WordファイルからExcelファイルへの変換処理のUIを定義したクラス
    /// </remarks>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 変換処理のキャンセルを制御するためのトークンソース
        /// </summary>
        private CancellationTokenSource m_cts;


        /// <summary>
        /// コンストラクタ。
        /// UI初期化およびコンバータのコールバック設定を行います。
        /// </summary>
        public MainWindow()
        {
            /* UIを初期化 */
            InitializeComponent();

            /* コールバックを設定 */
            W2EConverter.onProgressUpdate = this.UpdateProgressBar;
            W2EConverter.onLogUpdate = this.UpdateLog;
            W2MdConverter.onProgressUpdate = this.UpdateProgressBar;
            W2MdConverter.onLogUpdate = this.UpdateLog;
        }


        /// <summary>
        /// ウィンドウが表示された時の処理
        /// </summary>
        /// <remarks>
        /// 前回終了時の設定値を復元する
        /// </remarks>
        /// <param name="a_sender">イベント発生元オブジェクト</param>
        /// <param name="a_args">イベントデータ</param>
        private void WindowLoaded( object a_sender, RoutedEventArgs a_args )
        {
            /* 前回終了時の設定を復元する */
            this.m_wordPath.Text = Properties.Settings.Default.wordPath;
            this.m_markDown.IsChecked = ( this.m_markDown.Content.ToString() == Properties.Settings.Default.output );
            this.m_excel.IsChecked = !this.m_markDown.IsChecked.Value;
            this.EnableBtnConvert();
        }


        /// <summary>
        /// ウィンドウが閉じられた時の処理
        /// </summary>
        /// <remarks>
        /// 実行中の処理をキャンセルし、設定値を保存する
        /// </remarks>
        /// <param name="a_sender">イベント発生元オブジェクト</param>
        /// <param name="a_args">イベントデータ</param>
        private void WindowClosed( object a_sender, EventArgs a_args )
        {
            /* 実行中の処理をキャンセル */
            this.m_cts?.Cancel();

            /* 終了時の設定値を保存する */
            Properties.Settings.Default.wordPath = this.m_wordPath.Text;
            Properties.Settings.Default.output = ( this.m_excel.IsChecked.Value ) ? this.m_excel.Content.ToString() : this.m_markDown.Content.ToString();
            Properties.Settings.Default.Save();
        }


        /// <summary>
        /// 「...」ボタン押下時の処理
        /// </summary>
        /// <remarks>
        /// ファイル選択ダイアログを表示し、選択されたWordファイルのパスを設定する
        /// </remarks>
        /// <param name="a_sender">イベント発生元オブジェクト</param>
        /// <param name="a_args">イベントデータ</param>
        private void BtnOpenFileClick( object a_sender, RoutedEventArgs a_args )
        {
            var openFileDialog = new OpenFileDialog();
            if( openFileDialog.ShowDialog().Value )
            {
                this.m_wordPath.Text  = openFileDialog.FileName;
            }
        }


        /// <summary>
        /// Wordファイルパス変更時の処理
        /// </summary>
        /// <remarks>
        /// ファイルの存在有無に応じて変換ボタンの有効／無効を切り替える
        /// </remarks>
        /// <param name="a_sender">イベント発生元オブジェクト</param>
        /// <param name="a_args">イベントデータ</param>
        private void WordPathChanged( object sender, TextChangedEventArgs a_args )
        {
            this.EnableBtnConvert();
        }


        /// <summary>
        /// 「変換実行」ボタン押下時の処理
        /// </summary>
        /// <remarks>
        /// Word→Excel変換を実行する
        /// 変換完了後、自動的にExcelファイルを既定アプリで開く
        /// </remarks>
        /// <param name="a_sender">イベント発生元オブジェクト</param>
        /// <param name="a_args">イベントデータ</param>
        private async void BtnConvertClick( object a_sender, RoutedEventArgs a_args )
        {
            this.m_btnConvert.IsEnabled = false;
            this.m_btnCancel.IsEnabled = true;
            this.m_cts = new CancellationTokenSource();

            bool output2Excel_flg = this.m_excel.IsChecked.Value;
            string wordPath = this.m_wordPath.Text;
            string excelPath = Path.Combine(
                    Path.GetDirectoryName( wordPath ), 
                    Path.GetFileNameWithoutExtension( wordPath ) + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + ".xlsx" );
            string mdDir = Path.Combine(
                    Path.GetDirectoryName( wordPath), Path.GetFileName( wordPath ) + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) );

            /* ログ表示エリアを初期化 */
            this.m_log.Clear();

            try
            {
                await Task.Run( () =>
                {

                    /* 開始ログ */
                    this.UpdateLog( Environment.NewLine + "Wordファイル読込中..." + Environment.NewLine );

                    if( output2Excel_flg )
                    {
                        W2EConverter.Convert( wordPath, excelPath, this.m_cts.Token );
                    }
                    else
                    {
                        Directory.CreateDirectory( mdDir );
                        W2MdConverter.Convert( wordPath, mdDir, this.m_cts.Token );
                    }

                    /* 完了ログ */
                    this.UpdateLog( Environment.NewLine + "===== 変換完了 =====" );

                    /* ExcelファイルまたはMarkDownファイル出力フォルダを開く */
                    try
                    {
                        System.Diagnostics.Process.Start(
                            new System.Diagnostics.ProcessStartInfo()
                            {
                                FileName = output2Excel_flg ? excelPath : mdDir,
                                UseShellExecute = true
                            } );
                    }
                    catch( Exception ex )
                    {
                        Console.WriteLine( ex.Message );
                        UpdateLog( ex.Message );
                    }
                } );
            }
            finally
            {
                this.m_btnConvert.IsEnabled = true;
                this.m_btnCancel.IsEnabled = false;
            }
        }


        /// <summary>
        /// 「処理中断」ボタン押下時の処理
        /// </summary>
        /// <remarks>
        /// 実行中の変換処理にキャンセル要求を送信する
        /// </remarks>
        /// <param name="a_sender">イベント発生元オブジェクト</param>
        /// <param name="a_args">イベントデータ</param>
        private void BtnCancelClick( object a_sender, RoutedEventArgs a_args )
        {
            this.m_cts?.Cancel();
        }


        /// <summary>
        /// 変換ボタンの有効／無効を制御する
        /// </summary>
        /// <remarks>
        /// 指定されたパスにファイルが存在する場合のみ有効化する
        /// </remarks>
        private void EnableBtnConvert()
        {
            if( File.Exists( this.m_wordPath.Text ) )
            {
                this.m_btnConvert.IsEnabled = true;
            }
            else
            {
                this.m_btnConvert.IsEnabled = false;
            }
        }


        /// <summary>
        /// 進捗バーの値を更新する
        /// </summary>
        /// <remarks>
        /// バックグラウンドスレッドからの呼び出しに対応しています
        /// </remarks>
        /// <param name="a_value">進捗値</param>
        private void UpdateProgressBar( int a_value )
        {
            Dispatcher.Invoke( () =>
            {
                this.m_progressBar.Value = a_value;
            } );
        }


        /// <summary>
        /// ログ出力を更新する
        /// </summary>
        /// <remarks>
        /// バックグラウンドスレッドからの呼び出しに対応しています
        /// </remarks>
        /// <param name="a_value">出力するログメッセージ</param>
        private void UpdateLog( string a_value )
        {
            Dispatcher.Invoke( () =>
            {
                Console.WriteLine( a_value );
                this.m_log.AppendText( a_value + Environment.NewLine );
                this.m_log.ScrollToEnd();
            } );
        }


    }
}
