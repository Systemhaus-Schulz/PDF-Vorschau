using System;
using System.Windows;
using System.Windows.Threading;

namespace PDF_Vorschau
{
    public partial class SplashWindow : Window
    {
        private readonly DispatcherTimer _timer;
        private int _progress = 0;

        public SplashWindow()
        {
            InitializeComponent();

            _timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(50)
            };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            _progress += 2;
            if (_progress > 100)
                _progress = 100;

            ProgressBar.Value = _progress;

            if (_progress >= 100)
            {
                _timer.Stop();
                ShowMainWindow();
            }
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            _timer.Stop();
            ShowMainWindow();
        }

        private void ShowMainWindow()
        {
            var main = new MainWindow();
            main.Show();
            Close();
        }
    }
}
