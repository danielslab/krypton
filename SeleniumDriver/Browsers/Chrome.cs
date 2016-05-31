using System;
using System.Collections.Generic;
using System.IO;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using Common;
using System.Diagnostics;
using Driver.Browsers;

namespace Driver
{
    public class Chrome
    {
        private static int signal = 0;
        private static readonly ChromeOptions ChromeOpt = new ChromeOptions();
        private static readonly string ChromecookPath = Property.ApplicationPath + "ChromeCookies";
        internal static string ChromeProfilePath = string.Empty;

        public void IntializeDriver(ref IWebDriver driver, ref bool isBrowserDimendion, ref List<IWebDriver> driverlist,ref int width, ref int height) 
        {
            if (signal == 0)
            {
                if ((Directory.Exists(ChromeProfilePath)))
                {
                    string chromeProfileSourcePath = ChromeProfilePath;
                    DirectoryInfo chromeProfileSourceDir = new DirectoryInfo(chromeProfileSourcePath);

                    if (!chromeProfileSourceDir.Exists)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0012"));
                    if (File.Exists(ChromeProfilePath + @"\ChromeOptions.txt"))
                    {
                        using (StreamReader reader = new StreamReader(ChromeProfilePath + @"\ChromeOptions.txt"))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                ChromeOpt.AddArgument(line);
                            }
                        }
                    }
                }
                ChromeOpt.AddArguments("--test-type");
                ChromeOpt.AddArgument("--ignore-certificate-errors");
                ChromeOpt.AddArgument("--start-maximized");
                DirectoryInfo directory = new DirectoryInfo(ChromecookPath);
                Empty(directory);
                ChromeOpt.AddArguments("user-data-dir=" + ChromecookPath);
                Browser.Signal = 1;
            }
            driver = new ChromeDriver(Property.ApplicationPath + @"\Exes", ChromeOpt);
            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
            driver.Manage().Window.Maximize();
            if (isBrowserDimendion)
                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
            driverlist.Add(driver);
        }

        public static void Empty(DirectoryInfo directory)
        {
            try
            {
                foreach (FileInfo file in directory.GetFiles()) file.Delete();
                foreach (DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private Boolean ChromDriverIsRunning()
        {
            foreach (Process proc in Process.GetProcesses())
            {
                if (proc.ProcessName.ToString().ToLower() == KryptonConstants.CHROME_DRIVER)
                    return true;

            }
            return false;
        }

    }
}
