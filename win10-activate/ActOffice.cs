using System;
using System.Diagnostics;
using System.Windows;

using Microsoft.Win32;

namespace kms_activate
{
    class ActOffice
    {
        public static MainWindow mainW = (MainWindow)Application.Current.MainWindow;
        public static void OfficeActivate()
        {
            // debug info
            string kmsServerDbg, activateDbg;

            // make vol
            Process makeVol = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                WorkingDirectory = @"C:\",
                Arguments = @"//Nologo ospp.vbs /sethst:kms.jm33.me",

                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            // change KMS server
            mainW.Dispatcher.Invoke(() =>
            {
                startInfo.WorkingDirectory = mainW.OsppPath.Text;
                startInfo.Arguments = "//Nologo ospp.vbs /sethst:" + mainW.TextBox.Text;
            });
            Process kmsServer = new Process
            {
                StartInfo = startInfo
            };
            try
            {
                kmsServer.Start();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Exception caught", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            kmsServerDbg = kmsServer.StandardOutput.ReadToEnd();
            kmsServer.WaitForExit();

            // apply
            startInfo.Arguments = "//Nologo ospp.vbs /act";
            Process activate = new Process
            {
                StartInfo = startInfo
            };
            activate.Start();
            activateDbg = activate.StandardOutput.ReadToEnd();
            activate.WaitForExit();

            // display debug info
            mainW.Dispatcher.Invoke(() =>
            {
                if (mainW.ShowDebug.IsChecked == true)
                {
                    MessageBox.Show(kmsServerDbg + "\n" + activateDbg,
                        "Debug",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);
                }
            });

            // Check Office activation
            mainW.Dispatcher.Invoke(() =>
            {
                if (activateDbg.Contains("activation successful"))
                {
                    mainW.button.Content = "Done! Click to exit";
                    mainW.button.IsEnabled = true;
                    mainW.windows_option.IsEnabled = true;
                    mainW.office_option.IsEnabled = true;
                    mainW.ShowDebug.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("Activation failed!", "Failed!", MessageBoxButton.OK, MessageBoxImage.Error);
                    mainW.button.Content = "Retry";
                    mainW.button.IsEnabled = true;
                    mainW.windows_option.IsEnabled = true;
                    mainW.office_option.IsEnabled = true;
                    mainW.ShowDebug.IsEnabled = true;
                }
            });
        }

        public static void OfficeEnv()
        {

            // look for Office's install path, where OSPP.VBS can be found
            try
            {
                mainW.button.Content = "Activate ";
                RegistryKey localKey;
                if (Environment.Is64BitOperatingSystem)
                {
                    localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                }
                else
                {
                    localKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
                }

                string officepath = "";
                RegistryKey officeBaseKey = localKey.OpenSubKey(@"SOFTWARE\Microsoft\Office");
                if (officeBaseKey.OpenSubKey(@"16.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"16.0\Word\InstallRoot").GetValue("Path").ToString();

                    if (officepath.Contains("root"))
                    // Office 2019 can only be installed via Click-To-Run, therefore we get "C:\Program Files\Microsoft Office\root\Office16\",
                    // otherwise we get "C:\Program Files\Microsoft Office\Office16\"
                    {
                        // OSPP.VBS is still in "C:\Program Files\Microsoft Office\Office16\"
                        officepath = officepath.Replace("root", "");
                        mainW.button.Content += "Office 2019/2016";
                    }
                    else
                    {
                        mainW.button.Content += "Office 2016";
                    }
                }
                else if (officeBaseKey.OpenSubKey(@"15.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"15.0\Word\InstallRoot").GetValue("Path").ToString();
                    mainW.button.Content += "Office 2013";
                }
                else if (officeBaseKey.OpenSubKey(@"14.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"14.0\Word\InstallRoot").GetValue("Path").ToString();
                    mainW.button.Content += "Office 2010";
                }
                else
                {
                    MessageBox.Show("Only works with Office 2010 and/or above", "Unsupported version", MessageBoxButton.OK, MessageBoxImage.Error);
                    mainW.button.Content = "Unsupported version";
                    mainW.windows_option.IsChecked = true;
                    return;
                }
                mainW.OsppPath.Text = officepath;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Error detecting Office path", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
