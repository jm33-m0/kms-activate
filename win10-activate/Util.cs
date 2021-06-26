using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Threading;
using System.Windows;

namespace kms_activate
{
    class Util
    {
        public static MainWindow mainW = (MainWindow)Application.Current.MainWindow;
        public static bool IsAdmin()
        {
            bool isElevated;
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            isElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
            return isElevated;
        }

        public static bool IsActivated()
        {
            string status = RunProcess("cscript.exe", @"//NoLogo slmgr.vbs /dli", "", true);

            if (status.Contains("license status: licensed") || status.Contains("已授权"))
            {
                return true;
            }

            return false;
        }

        public static bool YesNo(string prompt, string title)
        {
            MessageBoxResult response = MessageBox.Show(prompt, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
            return response == MessageBoxResult.Yes;
        }

        public static string RunProcess(string name, string args, string workdir, bool silent)
        {
            ProcessStartInfo procInfo = new ProcessStartInfo
            {
                FileName = name,
                WorkingDirectory = System.Environment.GetEnvironmentVariable("SystemRoot") + @"\System32",
                Arguments = args,
            };
            if (workdir != "")
            {
                procInfo.WorkingDirectory = workdir;
            }
            if (silent)
            {
                procInfo.UseShellExecute = false;
                procInfo.CreateNoWindow = true;
                procInfo.WindowStyle = ProcessWindowStyle.Hidden;
                procInfo.RedirectStandardOutput = true;
            }

            Process proc = new Process
            {
                StartInfo = procInfo
            };
            try
            {
                proc.Start();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Exception caught", MessageBoxButton.OK, MessageBoxImage.Error);
                return "";
            }

            string output = proc.StandardOutput.ReadToEnd().ToLower();
            return output;
        }

        public static void CheckActivateState()
        {
            if (Util.IsActivated())
            {
                mainW.Dispatcher.Invoke(() =>
                {
                    mainW.button.Content = "Done! Click to exit";
                    mainW.button.IsEnabled = true;
                });
            }
            else
            {
                mainW.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("Activation failed!", "Failed!", MessageBoxButton.OK, MessageBoxImage.Error);
                    mainW.button.Content = "Retry";
                    mainW.button.IsEnabled = true;
                    mainW.windows_option.IsEnabled = true;
                    mainW.office_option.IsEnabled = true;
                    mainW.ShowDebug.IsEnabled = true;
                });
            }
        }

        public static void KMSActivate()
        {
            if (mainW.windows_option.IsChecked == true)
            {
                Thread win = new Thread(() =>
                {
                    ActWin.WinActivate();
                    CheckActivateState();
                });
                win.Start();
            }
            else if (mainW.office_option.IsChecked == true)
            {
                Thread office = new Thread(() =>
                {
                    ActOffice.OfficeActivate();
                });
                office.Start();
            }
        }

    }
}
