using System.Reflection;
using System.Windows;

namespace miccedux
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Title = Assembly.GetExecutingAssembly().GetName().Name.ToString();
            version.Content += Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Class1.SetWindow(this);
            Class1.SetFrame(this.MyFrame);
            MyFrame.Content = new MainPage();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (Class1.ExcelAlive)
            {
                e.Cancel = true;
                MessageBox.Show("Дождитесь окончания формирования документа во избежание появления процессов - зомби Excel");
            }
        }
    }
}
