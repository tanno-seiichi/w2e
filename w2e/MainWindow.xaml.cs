using Microsoft.Win32;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace w2e
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private CancellationTokenSource m_cts;

        public MainWindow()
        {
            InitializeComponent();
            W2EConverter.onProgressUpdate = this.UpdateProgressBar;
            W2EConverter.onLogUpdate = this.UpdateLog;
        }

        private void WindowLoaded( object a_sender, RoutedEventArgs a_args )
        {
            this.EnableBtnConvert();
        }

        private void WindowClosed( object a_sender, EventArgs a_args )
        {
            this.m_cts?.Cancel();
        }

        private void BtnOpenFileClick( object a_sender, RoutedEventArgs a_args )
        {
            var openFileDialog = new OpenFileDialog();
            if( openFileDialog.ShowDialog().Value )
            {
                this.m_wordPath.Text  = openFileDialog.FileName;
            }
        }

        private void WordPathChanged( object sender, System.Windows.Controls.TextChangedEventArgs e )
        {
            this.EnableBtnConvert();
        }

        private async void BtnConvertClick( object a_sender, RoutedEventArgs a_args )
        {
            this.m_btnConvert.IsEnabled = false;
            this.m_btnCancel.IsEnabled = true;
            this.m_cts = new CancellationTokenSource();

            string wordPath = this.m_wordPath.Text;
            string excelPath = Path.Combine(
                Path.GetDirectoryName( wordPath ), 
                Path.GetFileNameWithoutExtension( wordPath ) + "_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + ".xlsx" );

            try
            {
                await Task.Run( () =>
                {
                    W2EConverter.Convert( wordPath, excelPath, this.m_cts.Token );
                } );
            }
            finally
            {
                this.m_btnConvert.IsEnabled = true;
                this.m_btnCancel.IsEnabled = false;
            }
        }

        private void BtnCancelClick( object a_sender, RoutedEventArgs a_args )
        {
            this.m_cts?.Cancel();
        }

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

        private void UpdateProgressBar( int a_value )
        {
            Dispatcher.Invoke( () =>
            {
                this.m_progressBar.Value = a_value;
            } );
        }

        private void UpdateLog( string a_value )
        {
            Dispatcher.Invoke( () =>
            {
                Console.WriteLine( a_value );
                this.m_log.AppendText( a_value );
                this.m_log.ScrollToEnd();
            } );
        }

    }
}
