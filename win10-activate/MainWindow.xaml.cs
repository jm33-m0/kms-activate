using System.Windows;

namespace kms_activate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            // Must be admin
            if (!Util.IsAdmin())
            {
                MessageBox.Show("You have to perform this action as Admin!", "Privilege escalation required!", MessageBoxButton.OK, MessageBoxImage.Stop);
                Application.Current.Shutdown();
            }
            InitializeComponent();
            windows_option.IsChecked = true;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // when to exit
            if ((string)button.Content == "Done! Click to exit")
            {
                Application.Current.Shutdown();
            }

            if (office_option.IsChecked.Value)
            // Office has to be vol
            {
                MessageBoxResult response = MessageBox.Show("Make sure you are using VOL version", "Proceed?", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (response == MessageBoxResult.No)
                {
                    return;
                }
            }

            // Disable all buttons
            button.Content = "Please wait...";
            button.IsEnabled = false;
            windows_option.IsEnabled = false;
            office_option.IsEnabled = false;
            ShowDebug.IsEnabled = false;

            // Try to activate
            Util.KMSActivate();
        }

        private void Windows_option_Checked(object sender, RoutedEventArgs e)
        {
            ActWin.WinEnv();
        }

        private void Office_option_Checked(object sender, RoutedEventArgs e)
        {
            ActOffice.OfficeEnv();
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void TextBox_TextChanged_1(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}