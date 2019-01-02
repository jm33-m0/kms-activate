using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Principal;
using System.Threading;
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
            if (!(IsAdmin()))
            {
                MessageBox.Show("You have to perform this action as Admin!", "Privilege escalation required!", MessageBoxButton.OK, MessageBoxImage.Stop);
                Application.Current.Shutdown();
            }
            InitializeComponent();
            windows_option.IsChecked = true;
        }

        public bool IsAdmin()
        {
            bool isElevated;
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            isElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
            return isElevated;
        }

        public bool IsActivated(string choice)
        {
            ProcessStartInfo procInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                CreateNoWindow = true,
                UseShellExecute = false,
                WorkingDirectory = System.Environment.GetEnvironmentVariable("SystemRoot") + @"\System32",
                Arguments = @"//NoLogo slmgr.vbs /dli",
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true
            };

            Process licenseStatus = new Process
            {
                StartInfo = procInfo
            };
            try
            {
                licenseStatus.Start();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Exception caught", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            string status = licenseStatus.StandardOutput.ReadToEnd().ToLower();

            if (status.Contains("license status: licensed") || status.Contains("已授权"))
            {
                return true;
            }

            return false;
        }

        public void CheckActivateState(string choice)
        {
            if (IsActivated(choice))
            {
                this.Dispatcher.Invoke(() =>
                {
                    button.Content = "Done! Click to exit";
                    button.IsEnabled = true;
                });
            }
            else
            {
                this.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("Activation failed!", "Failed!", MessageBoxButton.OK, MessageBoxImage.Error);
                    button.Content = "Retry";
                    button.IsEnabled = true;
                    windows_option.IsEnabled = true;
                    office_option.IsEnabled = true;
                    CheckBox.IsEnabled = true;
                });
            }
        }

        public void KMSActivate()
        {
            if (windows_option.IsChecked == true)
            {
                Thread win = new Thread(() =>
                {
                    WinActivate();
                    CheckActivateState("win");
                });
                win.Start();
            }
            else if (office_option.IsChecked == true)
            {
                Thread office = new Thread(() =>
                {
                    OfficeActivate();
                    CheckActivateState("office");
                });
                office.Start();
            }
        }

        public void WinActivate()
        {
            // VOL keys
            Dictionary<string, string> winKeys = new Dictionary<string, string>()
            {
                {"Windows 10 Professional", "W269N-WFGWX-YVC9B-4J6C9-T83GX" },
                {"Windows 10 Enterprise", "NPPR9-FWDCX-D2C8J-H872K-2YT43" },
                {"Windows 8.1 Professional", "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9" },
                {"Windows 8.1 Enterprise", "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7" },
                {"Windows 7 Professional", "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4" },
                {"Windows 7 Enterprise", "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH" },
                {"Windows Server 2016 Standard", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" },
                {"Windows Server 2016 Datacenter", "CB7KF-BWN84-R7R2Y-793K2-8XDDG" },
                {"Windows Server 2012 R2 Server Standard", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX" },
                {"Windows Server 2012 R2 Datacenter", "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"},
                {"Windows Server 2008 R2 Standard", "YC6KT-GKW9T-YTKYR-T4X34-R7VHC" },
                {"Windows Server 2008 R2 Enterprise", "489J6-VHDMP-X63PK-3K798-CPX3Y" }
            };

            // which version to activate?
            string key = "";
            string productName = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "").ToString();
            if (productName.ToLower().Contains("ultimate"))
            // in case you want to activate Windows 7 Ultimate (it's a retail version, which doesn't support VOL at all)
            {
                MessageBox.Show("Not supported", "Error", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }


            try
            {
                foreach (string winversion in winKeys.Keys)
                {
                    if (winversion.Contains(productName))
                    {
                        key = winKeys[winversion];
                        this.Dispatcher.Invoke(() =>
                        {
                            button.Content = "Activating " + winversion + "...";
                        });
                        break;
                    }
                }
            }
            catch
            {
                MessageBox.Show("Windows version not supported", "Sorry", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }

            // debug info
            string makeVolDbg, kmsServerDbg, activateDbg;


            // make vol
            Process makeVol = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                WorkingDirectory = System.Environment.GetEnvironmentVariable("SystemRoot") + @"\System32",
                Arguments = @"//Nologo slmgr.vbs /ipk " + key,

                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            makeVol.StartInfo = startInfo;
            try
            {
                makeVol.Start();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Exception caught", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            makeVolDbg = makeVol.StandardOutput.ReadToEnd();
            makeVol.WaitForExit();

            // change KMS server
            this.Dispatcher.Invoke(() =>
            {
                startInfo.Arguments = "//Nologo slmgr.vbs /skms " + TextBox.Text;

            });
            Process kmsServer = new Process
            {
                StartInfo = startInfo
            };
            kmsServer.Start();
            kmsServerDbg = kmsServer.StandardOutput.ReadToEnd();
            kmsServer.WaitForExit();

            // apply
            startInfo.Arguments = "//Nologo slmgr.vbs /ato";
            Process activate = new Process
            {
                StartInfo = startInfo
            };
            activate.Start();
            activateDbg = activate.StandardOutput.ReadToEnd();
            activate.WaitForExit();

            // display debug info
            this.Dispatcher.Invoke(() =>
            {
                if (CheckBox.IsChecked == true)
                {
                    MessageBox.Show(makeVolDbg + "\n" + kmsServerDbg + "\n" + activateDbg,
                        "Debug",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);
                }
            });
        }

        public void OfficeActivate()
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
            this.Dispatcher.Invoke(() =>
            {
                startInfo.WorkingDirectory = OsppPath.Text;
                startInfo.Arguments = "//Nologo ospp.vbs /sethst:" + TextBox.Text;
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
            this.Dispatcher.Invoke(() =>
            {
                if (CheckBox.IsChecked == true)
                {
                    MessageBox.Show(kmsServerDbg + "\n" + activateDbg,
                        "Debug",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);
                }
            });

            // Check Office activation
            this.Dispatcher.Invoke(() =>
            {
                if (activateDbg.Contains("activation successful"))
                {
                    button.Content = "Done! Click to exit";
                    button.IsEnabled = true;
                    windows_option.IsEnabled = true;
                    office_option.IsEnabled = true;
                    CheckBox.IsEnabled = true;
                }
                else
                {
                    MessageBox.Show("Activation failed!", "Failed!", MessageBoxButton.OK, MessageBoxImage.Error);
                    button.Content = "Retry";
                    button.IsEnabled = true;
                    windows_option.IsEnabled = true;
                    office_option.IsEnabled = true;
                    CheckBox.IsEnabled = true;
                }
            });
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
            CheckBox.IsEnabled = false;

            // Try to activate
            KMSActivate();
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

        private void Windows_option_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                string winVersion = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "").ToString();
                button.Content = "Activate " + winVersion;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Error detecting Windows version", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Office_option_Checked(object sender, RoutedEventArgs e)
        {
            // look for Office's install path, where OSPP.VBS can be found
            try
            {
                button.Content = "Activate ";
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
                        button.Content += "Office 2019/2016";
                    }
                    else
                    {
                        button.Content += "Office 2016";
                    }
                }
                else if (officeBaseKey.OpenSubKey(@"15.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"15.0\Word\InstallRoot").GetValue("Path").ToString();
                    button.Content += "Office 2013";
                }
                else if (officeBaseKey.OpenSubKey(@"14.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"14.0\Word\InstallRoot").GetValue("Path").ToString();
                    button.Content += "Office 2010";
                }
                else
                {
                    MessageBox.Show("Only works with Office 2010 and/or above", "Unsupported version", MessageBoxButton.OK, MessageBoxImage.Error);
                    button.Content = "Unsupported version";
                    windows_option.IsChecked = true;
                    return;
                }
                OsppPath.Text = officepath;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Error detecting Office path", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}