using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

namespace Driver.Browsers
{
    class FireFox
    {
        internal static string FirefoxProfilePath = string.Empty;
        internal static FirefoxProfile FfProfile = null;
        DirectoryInfo _ffProfileDestDir = null;

        public void IntiailizeDriver(ref IWebDriver driver, ref bool isBrowserDimendion, ref List<IWebDriver> driverlist, ref int width, ref int height, ref string addonsPath)
        {
            if ((FirefoxProfilePath.Length != 0) && Directory.Exists(FirefoxProfilePath) && (FirefoxProfilePath.Split('.').Last().ToString().ToLower() != "zip"))
            {
                string ffProfileSourcePath = FirefoxProfilePath;
                DirectoryInfo ffProfileSourceDir = new DirectoryInfo(ffProfileSourcePath);

                if (!ffProfileSourceDir.Exists)
                    throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0012"));

                FfProfile = new FirefoxProfile(ffProfileSourceDir.ToString());
                if ((Common.Utility.GetParameter("FirefoxProfilePath")).Equals(Common.Utility.GetVariable("FirefoxProfilePath")))
                {
                    InitFireFox(ref driver,ref isBrowserDimendion,ref driverlist,ref width,ref height);
                    Common.Utility.SetVariable("FirefoxProfilePath", _ffProfileDestDir.ToString());
                }
                else
                {
                    //If Profile path is empty in parameter file then add setpreference to profile.: 
                    if (Common.Utility.GetParameter("FirefoxProfilePath").Equals(""))
                    {
                        FfProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
                    }
                    InitFireFox(ref driver, ref isBrowserDimendion, ref driverlist, ref width, ref height);
                }
            }
            else
            {
                // If "addonsPath" contains path then load addons from that directory path.
                if (addonsPath.Length != 0)
                {
                    try
                    {
                        string adPath = addonsPath;
                        if (!Directory.Exists(adPath))
                            throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0014"));
                        DirectoryInfo adDir = new DirectoryInfo(adPath);
                        DirectoryInfo adDirTemp = new DirectoryInfo("AddonsTemp");
                        adDirTemp.Create();
                        foreach (FileInfo fi in adDir.GetFiles())
                        {
                            fi.CopyTo(Path.Combine(adDirTemp.FullName, fi.Name), true);
                        }

                        int firebugCount = 0;
                        string[] addonFilePaths = Directory.GetFiles(adPath);
                        FfProfile = new FirefoxProfile();
                        FfProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
                        FfProfile.SetPreference("network.cookie.lifetimePolicy", 0);
                        FfProfile.SetPreference("privacy.sanitize.sanitizeOnShutdown", false);
                        FfProfile.SetPreference("privacy.sanitize.promptOnSanitize", false);
                        SetCommonPreferences();
                        //Assigning profile location so that firefox can next time use same profile to open
                        Common.Utility.SetVariable("FirefoxProfilePath", FfProfile.ProfileDirectory.ToString());
                        // Add all the addons found in the specified addon directory
                        foreach (FileInfo afi in adDirTemp.GetFiles())
                        {
                            if (afi.Name.ToLower().Contains("firebug"))
                                firebugCount++;
                            FfProfile.AddExtension(adDirTemp.Name + '/' + afi.Name);
                        }

                        // If user placed multiple versions of firebug in Addons directory then give error
                        if (firebugCount > 1)
                            throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0015"));

                        // Start firefox with the profile that contains all the addons in addon directory
                        driver = new FirefoxDriver(FfProfile);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                        driver.Manage().Window.Maximize();
                        if (isBrowserDimendion)
                            driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                        driverlist.Add(driver);
                    }
                    catch (Exception e)
                    {
                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0009").Replace("{MSG}", e.Message));
                    }
                }
                else
                {
                    int webDriverPort = 7055;
                    FfProfile = new FirefoxProfile();
                    FfProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
                    FfProfile.SetPreference("network.cookie.lifetimePolicy", 0);
                    FfProfile.SetPreference("privacy.sanitize.sanitizeOnShutdown", false);
                    FfProfile.SetPreference("privacy.sanitize.promptOnSanitize", false);
                    SetCommonPreferences();
                    for (int port = webDriverPort; port <= webDriverPort + 5; port++)
                    {
                        //Allowing security exception to appear and being able to handle.
                        FfProfile.SetPreference("browser.xul.error_pages.expert_bad_cert", true);
                        FfProfile.AcceptUntrustedCertificates = true;
                        string profileDir = FfProfile.ProfileDirectory;
                        FfProfile.Port = port;
                        driver = new FirefoxDriver(FfProfile);
                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                        driver.Manage().Window.Maximize();
                        if (isBrowserDimendion)
                            driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                        driverlist.Add(driver);
                        //Assigning profile location so that firefox can next time use same profile to open
                        Common.Utility.SetVariable("FirefoxProfilePath", FfProfile.ProfileDirectory.ToString());
                        break;
                    }

                    if (driver == null)
                    {
                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0016"));
                    }
                }
                //Add WebDriver Profile File to list of files that need to be deleted in temp folder.
                Common.Property.ListOfFilesInTempFolder.Add(Common.Utility.GetVariable("FirefoxProfilePath"));
            }
        }

        /// <summary>
        ///  Method to Intialize the FireFox Driver.
        /// <param name="driver">IWebDriver : ref of Driver variable to intialize it.</param>
        /// <param name="IsBrowserDimension">bool : Check the window size of the browser</param>
        /// <param name="driverlist">List: Maintain the List of Existing driver.</param>
        /// <param name="width">int : Use to set the size of the browser.</param>
        /// <param name="height">int : Set the height of the browser.</param>
        /// <returns></returns>
        /// </summary>
        private void InitFireFox(ref IWebDriver driver, ref bool IsBrowserDimendion, ref List<IWebDriver> driverlist, ref int width, ref int height)
        {
            //Support for Firefox version 5.0
            SetCommonPreferences();
            driver = new FirefoxDriver(FfProfile);
            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
            driver.Manage().Window.Maximize();
            if (IsBrowserDimendion)
                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
            driverlist.Add(driver);
            _ffProfileDestDir = new DirectoryInfo(FfProfile.ProfileDirectory);
            //Assigning profile location so that firefox can next time use same profile to open
        }

        /// <summary>
        /// Method to Set the Common Preferences of the firefox.
        /// </summary>
        private void SetCommonPreferences()
        {
            FfProfile.SetPreference("extensions.checkCompatibility.5.0", false);
            #region Handling Downloading Window.
            FfProfile.SetPreference("browser.download.dir", Common.Utility.GetParameter("downloadpath"));
            FfProfile.SetPreference("browser.download.useDownloadDir", true);
            FfProfile.SetPreference("browser.download.defaultFolder", Common.Utility.GetParameter("downloadpath"));
            FfProfile.SetPreference("browser.download.folderList", 2);
            FfProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/x-msdos-program, application/x-unknown-application-octet-stream, application/vnd.ms-powerpoint, application/excel, application/vnd.ms-publisher, application/x-unknown-message-rfc822, application/vnd.ms-excel, application/msword, application/x-mspublisher, application/x-tar, application/zip, application/x-gzip,application/x-stuffit,application/vnd.ms-works, application/powerpoint, application/rtf, application/postscript, application/x-gtar, video/quicktime, video/x-msvideo, video/mpeg, audio/x-wav, audio/x-midi, audio/x-aiff");
            #endregion
            #region Handling Unresponsive Script Warning.
            FfProfile.SetPreference("dom.max_script_run_time", 10 * 60);
            FfProfile.SetPreference("dom.max_chrome_script_run_time", 10 * 60);
            #endregion
        }



    }
}
