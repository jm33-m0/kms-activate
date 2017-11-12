using System.Security.Principal;
using System.Windows;
using System.Diagnostics;
using System.Management;
using System.Threading;

namespace win10_activate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public bool IsAdmin()
        {
            bool isElevated;
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            isElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
            return isElevated;
        }

        public static bool IsWindowsActivated()
        {
            ManagementScope scope = new ManagementScope(@"\\" + System.Environment.MachineName + @"\root\cimv2");
            scope.Connect();

            SelectQuery searchQuery = new SelectQuery("SELECT * FROM SoftwareLicensingProduct WHERE ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' and LicenseStatus = 1");
            ManagementObjectSearcher searcherObj = new ManagementObjectSearcher(scope, searchQuery);
                using (ManagementObjectCollection obj = searcherObj.Get())
                {
                    return obj.Count > 0;
                }
        }

        public void Chk()
        {
            if (IsWindowsActivated())
            {
                MessageBox.Show("Windows 10 has been activated!", "Success!", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Activation failed!", "Failed!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void KmsActivate()
        {
            // make vol
            System.Diagnostics.Process makeVol = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo
            {

                FileName = "cmd.exe",
                Arguments = "/C slmgr /ipk \"W269N-WFGWX-YVC9B-4J6C9-T83GX\"",

                CreateNoWindow = true,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            makeVol.StartInfo = startInfo;
            makeVol.Start();

            // change KMS server
            startInfo.Arguments = "/C slmgr /skms kms.03k.org";
            Process kmsServer = new Process
            {
                StartInfo = startInfo
            };
            kmsServer.Start();
            
            // apply
            startInfo.Arguments = "/C slmgr /ato";
            Process activate = new Process
            {
                StartInfo = startInfo
            };
            activate.Start();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (!(IsAdmin()))
            {
                MessageBox.Show("You have to perform this action as Admin!", "Privilege escalation required!", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }
            MessageBoxResult response = MessageBox.Show("This tool is for Windows 10 Professional only", "Proceed?", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (response == MessageBoxResult.No)
            {
                return;
            }
            button.Content = "Please wait...";
            button.IsEnabled = false;
            KmsActivate();
            Thread thread = new Thread(() =>
            {
                Chk();
            });
            thread.Start();
        }
    }
}