using Microsoft.Win32;

using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace kms_activate
{
    class ActOffice
    {
        public static MainWindow mainW = (MainWindow)Application.Current.MainWindow;
        public static void OfficeActivate()
        {
            if (!OfficeEnv())
            {
                Application.Current.Shutdown();
            }
            MessageBox.Show("Make sure to open Office and agree to user terms", "Tips", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.OK);

            // debug info
            string kmsServerDbg, activateDbg;

            // ospp root
            string root = "";

            // procinfo
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                WorkingDirectory = @"C:\",

                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            // change KMS server
            mainW.Dispatcher.Invoke(() =>
            {
                root = mainW.OsppPath.Text;
                startInfo.WorkingDirectory = mainW.OsppPath.Text;
                startInfo.Arguments = "//Nologo ospp.vbs /sethst:" + mainW.TextBox.Text;
            });

            if (IsOfficeActivated(root))
            {
                return;
            }

            Retail2Vol(root);

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
                IsOfficeActivated(startInfo.WorkingDirectory);
            });
        }

        public static bool IsOfficeActivated(string root)
        {
            // make vol
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "cscript.exe",
                WorkingDirectory = root,
                Arguments = @"//Nologo ospp.vbs /dstatus",

                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            // check license status
            Process activate = new Process
            {
                StartInfo = startInfo
            };
            activate.Start();
            string activateDbg = activate.StandardOutput.ReadToEnd();
            activate.WaitForExit();

            // Check Office activation
            mainW.Dispatcher.Invoke(() =>
            {
                if (mainW.ShowDebug.IsChecked == true)
                {
                    MessageBox.Show("Checking if Office is activated:\n" + activateDbg,
                        "Debug",
                        MessageBoxButton.OK,
                        MessageBoxImage.Asterisk);
                }

                if (activateDbg.Contains("activation successful") ||
                activateDbg.Contains("0xC004F009") ||
                activateDbg.Contains("LICENSE STATUS:  ---LICENSED---"))
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
            return false;
        }

        public static void Retail2Vol(string installRoot)
        /*
         * try to install vol key and licenses
         * product key is of Office Pro Plus
         */
        {
            string dstatus;
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "cscript.exe",
                    WorkingDirectory = installRoot,
                    Arguments = @"//Nologo ospp.vbs /dstatus",

                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                // check license status
                Process checkLicense = new Process
                {
                    StartInfo = startInfo
                };
                checkLicense.Start();
                dstatus = checkLicense.StandardOutput.ReadToEnd();
                checkLicense.WaitForExit();
            }
            catch (Exception err)
            {
                MessageBox.Show("Retail2Vol\n" + err.ToString(), "Exception caught", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (dstatus.ToLower().Contains("volume"))
            {
                return;
            }

            if (!Util.YesNo("You are not using a VOL version, try converting?", "Retail2Vol"))
            {
                return;
            }
            string licenseDir = installRoot + @"..\root\License";
            string key, visioKey, version;

            // handle different versions
            if (installRoot.EndsWith("Office16"))
            {
                licenseDir += @"16\";
                if (dstatus.Contains("Office 19"))
                {
                    key = "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP";
                    visioKey = "9BGNQ-K37YR-RQHF2-38RQ3-7VCBB";
                    version = "2019";
                }
                else
                {
                    key = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99";
                    visioKey = "PD3PC-RHNGV-FXJ29-8JK7D-RJRJK";
                    version = "2016";
                }
            }
            else if (installRoot.EndsWith("Office15"))
            {
                licenseDir += @"15\";
                key = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT";
                visioKey = "C2FG9-N6J68-H8BTJ-BW3QX-RM3B3";
                version = "2013";
            }
            else if (installRoot.EndsWith("Office14"))
            {
                licenseDir += @"14\";
                key = "VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB";
                visioKey = "7MCW8-VRQVK-G677T-PDJCM-Q8TCP";
                version = "2010";
            }
            else
            {
                MessageBox.Show("No compatible Office version found, exit?", "Goodbye", MessageBoxButton.OK, MessageBoxImage.Stop);
                return;
            }

            // try to convert retail version to VOL
            string log;
            log = Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "ProPlusVL_KMS_Client-ppd.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "ProPlusVL_KMS_Client-ul.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "ProPlusVL_KMS_Client-ul-oob.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "VisioProVL_KMS_Client-ppd.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "VisioProVL_KMS_Client-ul.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "VisioProVL_KMS_Client-ul-oob.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inslic:" + licenseDir + "pkeyconfig-office.xrm-ms", installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inpkey:" + key, installRoot, false);
            log += Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /inpkey:" + visioKey, installRoot, false);

            // tell user what happened
            MessageBox.Show("Converting Office " + version + " retail version to volume version...\n" + log, "Note: You are NOT using volume version", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static void InstallOffice()
        {
            if (!File.Exists("setup.exe") || !File.Exists("office-proplus.xml"))
            {
                MessageBox.Show("Make sure setup.exe and office-proplus.xml exist in current directory", "Files missing",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Util.RunProcess("setup.exe", "/configure office-proplus.xml", "./", true);
        }

        public static bool OfficeEnv()
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
                    var k = officeBaseKey.OpenSubKey(@"16.0\Word\InstallRoot");
                    if (k == null)
                    {
                        throw new Exception("Office not installed");
                    }

                    var val = k.GetValue("Path");
                    if (val == null)
                    {
                        throw new Exception("Office installation corrupted");
                    }
                    officeBaseKey.Close();
                    officepath= val.ToString();
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
                    officeBaseKey.Close();
                    mainW.button.Content += "Office 2013";
                }
                else if (officeBaseKey.OpenSubKey(@"14.0", false) != null)
                {
                    officepath = officeBaseKey.OpenSubKey(@"14.0\Word\InstallRoot").GetValue("Path").ToString();
                    officeBaseKey.Close();
                    mainW.button.Content += "Office 2010";
                }
                else
                {
                    MessageBox.Show("Only works with Office 2010 and/or above", "Unsupported version", MessageBoxButton.OK, MessageBoxImage.Error);
                    mainW.button.Content = "Unsupported version";
                    mainW.windows_option.IsChecked = true;
                    return false;
                }
                mainW.OsppPath.Text = officepath;

                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("Office installation not detected:\n" + err.ToString(), "Error detecting Office path", MessageBoxButton.OK);
                if (Util.YesNo("Download and install Office 2021 with Office Deployment Tool?", "Install Office"))
                {
                    InstallOffice();
                }
                return false;
            }
        }
    }
}
