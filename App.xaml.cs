using System.Windows;

namespace PDF_Vorschau
{
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            var main = new MainWindow();
            main.Show();
        }
    }
}
