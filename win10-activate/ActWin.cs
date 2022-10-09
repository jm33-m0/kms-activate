using Microsoft.Win32;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;

namespace kms_activate
{
    class ActWin
    {
        public static MainWindow mainW = (MainWindow)Application.Current.MainWindow;
        public static void WinActivate()
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
                {"Windows Server 2022 Standard", "VDYBN-27WPP-V4HQT-9VMD4-VMK7H" },
                {"Windows Server 2022 Datacenter", "WX4NM-KYWYW-QJJR4-XV3QB-6VM33" },
                {"Windows Server 2019 Standard", "N69G4-B89J2-4G8F4-WWYCC-J464C" },
                {"Windows Server 2019 Datacenter", "WMDGN-G9PQG-XVVXX-R3X43-63DFG" },
                {"Windows Server 2016 Standard", "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" },
                {"Windows Server 2016 Datacenter", "CB7KF-BWN84-R7R2Y-793K2-8XDDG" },
                {"Windows Server 2012 R2 Server Standard", "D2N9P-3P6X9-2R39C-7RTCD-MDVJX" },
                {"Windows Server 2012 R2 Datacenter", "W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9"},
                {"Windows Server 2008 R2 Standard", "YC6KT-GKW9T-YTKYR-T4X34-R7VHC" },
                {"Windows Server 2008 R2 Datacenter", "74YFP-3QFB3-KQT8W-PMXWJ-7M648" },
                {"Windows Server 2008 R2 Enterprise", "489J6-VHDMP-X63PK-3K798-CPX3Y" }
            };

            // which version to activate?
            string key = "";
            string productName = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "").ToString();

            foreach (string winversion in winKeys.Keys)
            {
                if (winversion.Contains(productName) || productName.Contains(winversion))
                {
                    key = winKeys[winversion];
                    if (productName.ToLower().Contains("evaluation"))
                    {
                        string edition = winversion.Split(' ')[winversion.Split(' ').Length - 1];
                        string args = "/online /set-edition:Server" + edition + " /productkey:" + key + " /accepteula";
                        string eval2license = Util.RunProcess("dism.exe", args, "", false);
                        if (eval2license == "")
                        {
                            MessageBox.Show("Evaluation version failed to be converted", "Sorry", MessageBoxButton.OK, MessageBoxImage.Stop);
                            Application.Current.Shutdown();
                        }
                        MessageBox.Show("To upgrade to licensed version, a reboot is required\n" + args + "\n" + eval2license,
                            "Note", MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        Application.Current.Shutdown();
                    }
                    mainW.Dispatcher.Invoke(() =>
                    {
                        mainW.button.Content = "Activating " + winversion + "...";
                    });
                    break;
                }
            }

            if (key == "")
            {
                MessageBox.Show("Windows version not supported", "Sorry", MessageBoxButton.OK, MessageBoxImage.Stop);
                Application.Current.Shutdown();
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
            mainW.Dispatcher.Invoke(() =>
            {
                startInfo.Arguments = "//Nologo slmgr.vbs /skms " + mainW.TextBox.Text;

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
            mainW.Dispatcher.Invoke(() =>
            {
                if (mainW.ShowDebug.IsChecked == true)
                {
                    MessageBox.Show(makeVolDbg + "\n" + kmsServerDbg + "\n" + activateDbg,
                        "Debug",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);
                }
            });
        }

        public static void WinEnv()
        {
            try
            {
                string winVersion = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "").ToString();
                mainW.button.Content = "Activate " + winVersion;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.ToString(), "Error detecting Windows version", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
