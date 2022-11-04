using System.Windows;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Organizer.Classes.ClassFrame.frameMain = FraimMainWin;
            FraimMainWin.Navigate(new Pages.Calendar());
        }

        private void BtnGoPageСalendar_Click(object sender, RoutedEventArgs e)
        {
            FraimMainWin.Navigate(new Pages.Calendar());
        }

        private void BtnGoPageTemplate_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnGoSettings_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
