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
            // convert to VOL if needed
            Retail2Vol();

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

        public static void Retail2Vol()
        /*
         * try to install vol key and licenses
         * product key is of Office Pro Plus
         */
        {
            string installRoot = mainW.OsppPath.Text;
            string dstatus = Util.RunProcess("cscript.exe", "//NoLogo ospp.vbs /dstatus", installRoot, false);
            if (dstatus.ToLower().Contains("volume"))
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
