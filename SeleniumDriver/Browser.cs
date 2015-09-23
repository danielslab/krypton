/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Driver.Browser.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Browser action class
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Data;
using System.Windows.Forms.Design;
using OpenQA.Selenium.Remote;
using Selenium;
using Common;
using System.Net;
using OpenQA.Selenium.Safari;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using System.Diagnostics;
using System.ServiceProcess;


namespace Driver
{

    public class Browser
    {
       
        public static IWebDriver driver;
        public static string browserName;
        //public static string browserVersion;
        private Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        private static Selenium.DefaultSelenium selenium = null;
        private static string errorCaptureAs = string.Empty;
        public static string isRemoteExecution = string.Empty;
        private static string remoteUrl = string.Empty;
        private static string firefoxProfilePath = string.Empty;
        private static string chromeProfilePath = string.Empty;//Firefox profile path is get from parameters.ini
        private static string addonsPath = string.Empty;    // Addons path is get from parameters.ini
        private IAlert alert = null;
        public WebDriverWait alertLoadingWait; //used for asynchronous wait for alerts.
        public Func<IWebDriver, IAlert> delAlertLoaded;
        public string waitTimeForAlerts = string.Empty;
        private static Browser browser = null;
        private FirefoxProfile ffProfile = null;
        private static string driverName = string.Empty;
        private string username = string.Empty;
        private string accessKey = string.Empty;
        public static List<IWebDriver> driverlist = new List<IWebDriver>();  // to maintain all opened browser(used in close Browser method)
         public static int width = 600; // default values for Browser Window
        public static int height = 800;
        public  static Boolean IsBrowserDimendion = false;
        //Advanced user Interaction API object for driver
        public static Actions driverActions;
        private static ChromeOptions chromeOpt = new ChromeOptions();
        private static int signal = 0;
        private static string chromecookPath = Common.Property.ApplicationPath +"ChromeCookies";


        /// <summary>
        ///  Constructor to call initDriver method and set browser name variable.
        /// </summary>
        public Browser(string errorCapture, string browserName = null,bool DeleteCookie=true)
        {
            
            if (string.IsNullOrWhiteSpace(browserName) == false)
            {
                Browser.browserName = browserName;
                this.initDriver(DeleteCookie);
            }

            if (string.IsNullOrWhiteSpace(errorCapture) == false)
            {
                errorCaptureAs = errorCapture;
            }
            else
            {
                errorCapture = "image";
            }
        }


        public void Empty(System.IO.DirectoryInfo directory)
        {
            try
            {
                foreach (System.IO.FileInfo file in directory.GetFiles()) file.Delete();
                foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
        /// <summary>
        ///  set object information dictionary.
        /// </summary>
        /// <param name="objDataRow">Dictionary : Test Object information dictionary</param>
        public void SetObjDataRow(Dictionary<string, string> objDataRow)
        {
            this.objDataRow = objDataRow;
            // Refrencing waitLoading Object from TestObject.
            alertLoadingWait = TestObject.objloadingWait;
        }



        /// <summary>
        /// This method will provide testobject information and attribute. 
        /// </summary>
        /// <param name="dataType">string : Name of the test object</param>
        /// <returns >string :  object information</returns>
        private string GetData(string dataType)
        {
            try
            {
                switch (dataType)
                {
                    case "logical_name":
                        return objDataRow["logical_name"];
                    case KryptonConstants.OBJ_TYPE:
                        return objDataRow[KryptonConstants.OBJ_TYPE];
                    case KryptonConstants.HOW:
                        return objDataRow[KryptonConstants.HOW];
                    case KryptonConstants.WHAT:
                        return objDataRow[KryptonConstants.WHAT];
                    case KryptonConstants.MAPPING:
                        return objDataRow[KryptonConstants.MAPPING];
                    default:
                        return null;
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                throw new System.Collections.Generic.KeyNotFoundException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0008"));
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        /// Assigning platform during execution.
        /// </summary>
        /// <param name="platformString">string : platform string like Mac, Linux, Windows etc.</param>
        /// <returns>Platform Type Object.</returns>
        private PlatformType assignPlatform()
        {
            PlatformType platform;
            string platformString = string.Empty;
            try
            {
                platformString = Common.Utility.GetParameter("Platform");
            }
            catch { }
            switch (platformString.ToLower())
            {
                case "mac":
                    platform = PlatformType.Mac;
                    break;
                case "linux":
                    platform = PlatformType.Linux;
                    break;
                default:
                    platform = PlatformType.Windows;
                    break;
            }
            return platform;
        }


        /// <summary>
        ///  This method initialize specified web driver.
        ///  Updated for Window Resizing 
        /// /// </summary>
        private void initDriver(bool DeleteCookie=true)
        {
            try
            {

                Property.isSauceLabExecution = false;
                if (string.Compare(isRemoteExecution, "false", true) == 0)
                {
                    switch (browserName.ToLower())
                    {
                        case KryptonConstants.BROWSER_FIREFOX:
                            // If "firefoxProfilePath" contains path then load profile from that directory path.
                            if ((Browser.firefoxProfilePath.Length != 0) && Directory.Exists(Browser.firefoxProfilePath) && (Browser.firefoxProfilePath.Split('.').Last().ToString().ToLower() != "zip"))
                                
                            {
                                string ffProfileSourcePath = Browser.firefoxProfilePath;
                                DirectoryInfo ffProfileSourceDir = new DirectoryInfo(ffProfileSourcePath);

                                if (!ffProfileSourceDir.Exists)
                                    throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0012"));

                                if ((Common.Utility.GetParameter("FirefoxProfilePath")).Equals(Common.Utility.GetVariable("FirefoxProfilePath")))
                                {
                                    #region  This section works for 2.0rc2 release of selenium

                                    ffProfile = new FirefoxProfile(ffProfileSourceDir.ToString());
                                    //Support for Firefox version 5.0
                                    ffProfile.SetPreference("extensions.checkCompatibility.5.0", false);
                                    #region Handling Downloading Window.
                                    ffProfile.SetPreference("browser.download.dir", Common.Utility.GetParameter("downloadpath"));
                                    ffProfile.SetPreference("browser.download.useDownloadDir", true);
                                    ffProfile.SetPreference("browser.download.defaultFolder", Common.Utility.GetParameter("downloadpath"));
                                    ffProfile.SetPreference("browser.download.folderList", 2);
                                    ffProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/x-msdos-program, application/x-unknown-application-octet-stream, application/vnd.ms-powerpoint, application/excel, application/vnd.ms-publisher, application/x-unknown-message-rfc822, application/vnd.ms-excel, application/msword, application/x-mspublisher, application/x-tar, application/zip, application/x-gzip,application/x-stuffit,application/vnd.ms-works, application/powerpoint, application/rtf, application/postscript, application/x-gtar, video/quicktime, video/x-msvideo, video/mpeg, audio/x-wav, audio/x-midi, audio/x-aiff");
                                    #endregion
                                    #region Handling Unresponsive Script Warning.
                                    ffProfile.SetPreference("dom.max_script_run_time", 10 * 60);
                                    ffProfile.SetPreference("dom.max_chrome_script_run_time", 10 * 60);
                                    #endregion
                                    driver = new FirefoxDriver(ffProfile);
                                    driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                                    driver.Manage().Window.Maximize();
                                    if (IsBrowserDimendion)
                                        driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                                    driverlist.Add(driver);
                                    DirectoryInfo ffProfileDestDir = new DirectoryInfo(ffProfile.ProfileDirectory);

                                    //Assigning profile location so that firefox can next time use same profile to open
                                    Common.Utility.SetVariable("FirefoxProfilePath", ffProfileDestDir.ToString());
                                    #endregion
                                }

                               // This else loop approches where there is a profile already created within same test case
                                else
                                {
                                    ffProfile = new FirefoxProfile(ffProfileSourceDir.ToString());
                                    //If Profile path is empty in parameter file then add setpreference to profile.: 
                                    if (Common.Utility.GetParameter("FirefoxProfilePath").Equals(""))
                                    {
                                        ffProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
                                    }
                                    //For ff5.0 version support.
                                    ffProfile.SetPreference("extensions.checkCompatibility.5.0", false);
                                    #region Handling Downloading Window.
                                    ffProfile.SetPreference("browser.download.dir", Common.Utility.GetParameter("downloadpath"));
                                    ffProfile.SetPreference("browser.download.useDownloadDir", true);
                                    ffProfile.SetPreference("browser.download.defaultFolder", Common.Utility.GetParameter("downloadpath"));
                                    ffProfile.SetPreference("browser.download.folderList", 2);
                                    ffProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/x-msdos-program, application/x-unknown-application-octet-stream, application/vnd.ms-powerpoint, application/excel, application/vnd.ms-publisher, application/x-unknown-message-rfc822, application/vnd.ms-excel, application/msword, application/x-mspublisher, application/x-tar, application/zip, application/x-gzip,application/x-stuffit,application/vnd.ms-works, application/powerpoint, application/rtf, application/postscript, application/x-gtar, video/quicktime, video/x-msvideo, video/mpeg, audio/x-wav, audio/x-midi, audio/x-aiff");
                                    #endregion
                                    #region Handling Unresponsive Script Warning.
                                    ffProfile.SetPreference("dom.max_script_run_time", 10 * 60);
                                    ffProfile.SetPreference("dom.max_chrome_script_run_time", 10 * 60);
                                    #endregion
                                    ffProfile.Port = 9966;
                                    driver = new FirefoxDriver(ffProfile);
                                    driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                                    driver.Manage().Window.Maximize();
                                    if (IsBrowserDimendion)
                                        driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                                    driverlist.Add(driver);
                                }

                            }
                            else
                            {
                                // If "addonsPath" contains path then load addons from that directory path.
                                if (Browser.addonsPath.Length != 0)
                                {
                                    try
                                    {
                                        string adPath = Browser.addonsPath;

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
                                        ffProfile = new FirefoxProfile();
                                        ffProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
                                        ffProfile.SetPreference("network.cookie.lifetimePolicy", 0);
                                        ffProfile.SetPreference("privacy.sanitize.sanitizeOnShutdown", false);
                                        ffProfile.SetPreference("privacy.sanitize.promptOnSanitize", false);
                                        // for ff5 support.
                                        ffProfile.SetPreference("extensions.checkCompatibility.5.0", false);
                                        #region Handling Downloading Window.
                                        ffProfile.SetPreference("browser.download.dir", Common.Utility.GetParameter("downloadpath"));
                                        ffProfile.SetPreference("browser.download.useDownloadDir", true);
                                        ffProfile.SetPreference("browser.download.defaultFolder", Common.Utility.GetParameter("downloadpath"));
                                        ffProfile.SetPreference("browser.download.folderList", 2);
                                        ffProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/x-msdos-program, application/x-unknown-application-octet-stream, application/vnd.ms-powerpoint, application/excel, application/vnd.ms-publisher, application/x-unknown-message-rfc822, application/vnd.ms-excel, application/msword, application/x-mspublisher, application/x-tar, application/zip, application/x-gzip,application/x-stuffit,application/vnd.ms-works, application/powerpoint, application/rtf, application/postscript, application/x-gtar, video/quicktime, video/x-msvideo, video/mpeg, audio/x-wav, audio/x-midi, audio/x-aiff");
                                        #endregion
                                        #region Handling Unresponsive Script Warning.
                                        ffProfile.SetPreference("dom.max_script_run_time", 10 * 60);
                                        ffProfile.SetPreference("dom.max_chrome_script_run_time", 10 * 60);
                                        #endregion
                                        //Assigning profile location so that firefox can next time use same profile to open
                                        Common.Utility.SetVariable("FirefoxProfilePath", ffProfile.ProfileDirectory.ToString());

                                        // Add all the addons found in the specified addon directory
                                        foreach (FileInfo afi in adDirTemp.GetFiles())
                                        {
                                            if (afi.Name.ToLower().Contains("firebug"))
                                                firebugCount++;

                                            ffProfile.AddExtension(adDirTemp.Name + '/' + afi.Name);
                                        }

                                        // If user placed multiple versions of firebug in Addons directory then give error
                                        if (firebugCount > 1)
                                            throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0015"));

                                        // Start firefox with the profile that contains all the addons in addon directory
                                        driver = new FirefoxDriver(ffProfile);
                                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                                        driver.Manage().Window.Maximize();
                                        if (IsBrowserDimendion)
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
                                    ffProfile = new FirefoxProfile();
                                    ffProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
                                    ffProfile.SetPreference("network.cookie.lifetimePolicy", 0);
                                    ffProfile.SetPreference("privacy.sanitize.sanitizeOnShutdown", false);
                                    ffProfile.SetPreference("privacy.sanitize.promptOnSanitize", false);
                                    ffProfile.SetPreference("extensions.checkCompatibility.5.0", false);//For FF5 version support.

                                    #region Handling Downloading Window.
                                    ffProfile.SetPreference("browser.download.dir", Common.Utility.GetParameter("downloadpath"));
                                    ffProfile.SetPreference("browser.download.useDownloadDir", true);
                                    ffProfile.SetPreference("browser.download.defaultFolder", Common.Utility.GetParameter("downloadpath"));
                                    ffProfile.SetPreference("browser.download.folderList", 2);
                                    ffProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/x-msdos-program, application/x-unknown-application-octet-stream, application/vnd.ms-powerpoint, application/excel, application/vnd.ms-publisher, application/x-unknown-message-rfc822, application/vnd.ms-excel, application/msword, application/x-mspublisher, application/x-tar, application/zip, application/x-gzip,application/x-stuffit,application/vnd.ms-works, application/powerpoint, application/rtf, application/postscript, application/x-gtar, video/quicktime, video/x-msvideo, video/mpeg, audio/x-wav, audio/x-midi, audio/x-aiff");
                                    #endregion
                                    #region Handling Unresponsive Script Warning.
                                    ffProfile.SetPreference("dom.max_script_run_time", 10 * 60);
                                    ffProfile.SetPreference("dom.max_chrome_script_run_time", 10 * 60);
                                    #endregion
                                    for (int port = webDriverPort; port <= webDriverPort + 5; port++)
                                    {
                                        // Allowing security exception to appear and being able to handle
                                        ffProfile.SetPreference("browser.xul.error_pages.expert_bad_cert", true);
                                        ffProfile.AcceptUntrustedCertificates = true;
                                        string profileDir = ffProfile.ProfileDirectory;
                                        ffProfile.Port = port;
                                        driver = new FirefoxDriver(ffProfile);
                                        driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                                        driver.Manage().Window.Maximize();
                                        if (IsBrowserDimendion)
                                            driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                                        driverlist.Add(driver);
                                        //Assigning profile location so that firefox can next time use same profile to open
                                        Common.Utility.SetVariable("FirefoxProfilePath", ffProfile.ProfileDirectory.ToString());

                                        break;
                                    }

                                    if (driver == null)
                                    {
                                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0016"));
                                    }
                                }
                                //Add WebDriver Profile File to list of files that need to be deleted in temp folder.
                                Common.Property.listOfFilesInTempFolder.Add(Common.Utility.GetVariable("FirefoxProfilePath"));
                            }
                            break;
                        case KryptonConstants.BROWSER_IE:

                            InternetExplorerOptions options = new InternetExplorerOptions();
                            // Commented to check CSA related problem
                            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                         //for merging mobile build   
                            
                         options.EnablePersistentHover = false; //added for IE-8 certificate related issue.
                            if (DeleteCookie)
                            {
                                options.EnsureCleanSession = true;
                            }
                            driver = new InternetExplorerDriver(Property.ApplicationPath+@"\Exes",options);
                            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                            driver.Manage().Window.Maximize();    //  Added to maximize IE window forcibely, as this code is updated Action file.                                        
                            if (IsBrowserDimendion)
                                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                            driverlist.Add(driver);
                            break;

                        case Common.KryptonConstants.BROWSER_CHROME:
                            // Selenium v 2.1.17 states to use chrome options instead of capabilities
                            if(signal==0)
                            { 
                                if ((Directory.Exists(Browser.chromeProfilePath)))
                                {
                                    string chromeProfileSourcePath = Browser.chromeProfilePath;
                                    DirectoryInfo chromeProfileSourceDir = new DirectoryInfo(chromeProfileSourcePath);

                                    if (!chromeProfileSourceDir.Exists)
                                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0012"));
                                    if (File.Exists(chromeProfilePath+ @"\ChromeOptions.txt"))
                                    {
                                        using (StreamReader reader = new StreamReader(chromeProfilePath + @"\ChromeOptions.txt"))
                                        {
                                            string line = string.Empty;
                                            while ((line = reader.ReadLine()) != null)
                                            {
                                                chromeOpt.AddArgument(line);    
                                            }
                                            
                                        }
                                    }
                                }
                            chromeOpt.AddArguments("--test-type");
                            chromeOpt.AddArgument("--ignore-certificate-errors");
                            chromeOpt.AddArgument("--start-maximized");
                            System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(chromecookPath);
                            Empty(directory);
                            chromeOpt.AddArguments("user-data-dir="+chromecookPath);
                            signal = 1;
                            }
                            driver = new ChromeDriver(Property.ApplicationPath+@"\Exes", chromeOpt);
                            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                            driver.Manage().Window.Maximize();
                            if (IsBrowserDimendion)
                                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                            driverlist.Add(driver);
                            break;
                        case KryptonConstants.BROWSER_SAFARI:
                            driver = new SafariDriver();
                            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
                            driver.Manage().Window.Maximize();
                            if (IsBrowserDimendion)
                                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                            driverlist.Add(driver);
                            break;
                        default:
                            Console.WriteLine("No browser is defined.");
                            break;
                    }
                    WebDriverWait wdw = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, new TimeSpan(1000));
                }
                else
                {
                    // Attempt by  for sauce execution
                    //Start
                    if (Common.Property.RemoteUrl.ToLower().Contains("saucelabs"))
                    {
                        DesiredCapabilities capabilities = new DesiredCapabilities();
                            Utility.SetParameter("CloseBrowserOnCompletion", "true");//Forcing the close browser to true.
                            Utility.SetVariable("CloseBrowserOnCompletion", "true");
                            Property.isSauceLabExecution = true;
                            capabilities.SetCapability("username", Common.Utility.GetParameter("username"));//Registered user name of Sauce labs
                            capabilities.SetCapability("accessKey", Common.Utility.GetParameter("password"));// Accesskey provided by the Sauce labs 
                            capabilities.SetCapability("platform", Common.Utility.GetParameter("Platform"));// OS on which execution is to be done Eg: Windows 7 , mac , Linux etc..
                            capabilities.SetCapability("name", Common.Utility.GetParameter("TestCaseId"));
                            capabilities.SetCapability("browser", Common.Utility.GetParameter("SauceBrowser"));
                            capabilities.SetCapability("version", Common.Utility.GetParameter("VersionofBrowser"));
                            string RemoteHost = string.Empty;
                            remoteUrl = Common.Property.RemoteUrl + "/wd/hub";
                            
                            // if Sauce connect is required...
                            string isSauceConnectRequired = Common.Utility.GetParameter("IsTestEnvironment");
                            if (isSauceConnectRequired.ToLower()=="true")
                            {
                                executeSauceConnect();
                                Thread.Sleep(20 * 1000);
                            }

                            SeleniumGrid oSeleniumGrid = new SeleniumGrid(remoteUrl, Common.Utility.GetParameter("SauceBrowser"), capabilities);
                            Browser.driver = oSeleniumGrid.GetDriverSauce();
                            if (!string.IsNullOrEmpty(RemoteHost))
                            {
                                Common.Property.RCMachineId = RemoteHost;
                                Utility.SetVariable(Common.Property.RCMachineId, RemoteHost);
                                Utility.SetParameter(Common.Property.RCMachineId, RemoteHost);
                            }
                            driverActions = new Actions(driver);

                    }
                    else
                    {
                        string RemoteHost = string.Empty;
                    remoteUrl = Common.Property.RemoteUrl + "/wd/hub";
                     // (management of remote driver in a seperate class)
                    SeleniumGrid oSeleniumGrid = new SeleniumGrid(remoteUrl, browserName);
                    Browser.driver = oSeleniumGrid.GetDriver(out RemoteHost);
                    if (!string.IsNullOrEmpty(RemoteHost))
                    {
                        Common.Property.RCMachineId = RemoteHost;
                        Utility.SetVariable(Common.Property.RCMachineId, RemoteHost);
                        Utility.SetParameter(Common.Property.RCMachineId, RemoteHost);                  

                    }

                    //Initializing actions object for later usage
                    driverActions = new Actions(driver);
                    }     
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL) || e.Message.IndexOf("Connection refused", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (isRemoteExecution.Equals("true"))
                    {

                        this.initDriver();

                    }
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw e;
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf(Common.exceptions.ERROR_PARSINGVALUE, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (isRemoteExecution.Equals("true"))
                    {
                        this.initDriver();
                    }
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw e;
            }

        
        }

        /// <summary>
        /// Handle Certificate Navigation block in IE.
        /// </summary>
        public static void handleIECertificateError()
        {
            try
            {
                if (driver.Title.Contains("Certificate Error: Navigation Blocked"))
                {
                    driver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
                }
            }
            catch
            {
                //Nothing to throw.
            }

        }

        /// <summary>
        /// WaitAndGetAlert will wait until expected alert is returned OR specified time has been exceeded.
        /// </summary>
        /// <param name="ldriver">IwebDriver instance</param>
        /// <returns>IAlert : Alert Interface </returns>
        public IAlert waitAndGetAlert(IWebDriver ldriver)
        {
           
            double totalSeconds = (double)DateTime.Now.Ticks / TimeSpan.TicksPerSecond;

            while (((double)DateTime.Now.Ticks / TimeSpan.TicksPerSecond - totalSeconds) <= Common.Property.Waitforalert)
            {
                if (isAlertPresent())
                {
                    string text = driver.SwitchTo().Alert().Text;
                    return driver.SwitchTo().Alert();
                }

            }

            return null;

        }
        //This method is checking firefox.
        private Boolean FirefoxDriverIsRunning()
        {
            try
            {
                foreach (Process proc in System.Diagnostics.Process.GetProcesses())
                {
                    if (proc.ProcessName.ToString().ToLower() == "firefox" || proc.ProcessName.ToString().ToLower() == "firefoxdriver")
                        return true;

                }

                return false;
            }
            catch
            {

            }
            return true;
        }
        private Boolean ChromDriverIsRunning()
        {
            try
            {
                foreach (Process proc in System.Diagnostics.Process.GetProcesses())
                {
                    if (proc.ProcessName.ToString().ToLower() == KryptonConstants.CHROME_DRIVER)
                        return true;

                }

                return false;
            }
            catch
            {

            }
            return true;
        }

          /// <summary>
        /// Method to close current browser associate with web driver 
        /// </summary>
        public void CloseBrowser()//this method set arrKnownBrowserHwnd in TestObject class. 
        {
            try
            {
                try
                {
                    switchToMostRecentBrowser();
                }
                catch
                { }

                //back to previous different browser.
                try
               {

                   driver.Close();
                   Thread.Sleep(3000);
                    if (driverlist.Count > 1)
                    {
                        for (int i = driverlist.Count - 1; i >= 0; i--)
                        {
                            try
                            {
                                driver = driverlist[i];
                                string window = switchToMostRecentBrowser();
                                driver = driver.SwitchTo().Window(window);
                                driver.Title.ToString();  // Test for Exception with current driver.
                                break;
                            }
                            catch (Exception exx)
                            {
                                continue;
                            }

                        }
                        if (driver.ToString().IndexOf("InternetExplorer", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            browserName = KryptonConstants.BROWSER_IE;
                            driverName = "InternetExplorerDriver";

                        }
                        else if (driver.ToString().IndexOf(KryptonConstants.BROWSER_CHROME, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            driverName = KryptonConstants.CHROME_DRIVER;
                            browserName = KryptonConstants.BROWSER_CHROME;
                        }
                        else if (driver.ToString().IndexOf(KryptonConstants.BROWSER_FIREFOX, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            driverName = KryptonConstants.FIREFOX_DRIVER;
                            browserName = KryptonConstants.BROWSER_FIREFOX;
                        }
                        else if (driver.ToString().IndexOf(KryptonConstants.BROWSER_SAFARI, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            driverName = KryptonConstants.SAFARI_DRIVER;
                            browserName = KryptonConstants.BROWSER_SAFARI;
                        }
                        Driver.TestObject.driver = driver;
                        Driver.Browser.driver = Driver.TestObject.driver;
                    }
                }
                catch (Exception ex) 
                {
                    //Do nothing No browser to close.
                 }

                if (browserName.ToLower().Equals(KryptonConstants.BROWSER_FIREFOX) && Property.arrKnownBrowserHwnd.Count().Equals(1))
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    Func<IWebDriver, bool> condition = this.ProfileRunningCondition;
                    try
                    {
                        wait.Until(condition);
                    }
                    catch
                    {
                       
                    }
                }
                
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }

            catch (NullReferenceException e)
            {
                throw new NullReferenceException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0010") + ":" + e.Message);
            }

            catch (Exception e) 
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }

        private void updateJobDeatils()
        {
            try
            {
                HttpWebRequest request = WebRequest.Create("https://" + username + ":" + accessKey + "@saucelabs.com/rest/v1/" + username + "/jobs/" + Common.Utility.GetVariable("SessionID")) as HttpWebRequest;
                request.Method = "Put";
                request.Credentials = new NetworkCredential(username, accessKey);
                request.ContentType = "application/json";
                //Setting the status.
                string ByteString = "{\"passed\": true} ";
                if (Property.JobExecutionStatus.Equals(ExecutionStatus.Fail)) { ByteString = "{\"passed\": false}"; }
                byte[] byteRequestBody = Encoding.UTF8.GetBytes(ByteString);
                Stream requestStream = request.GetRequestStream();
                requestStream.Write(byteRequestBody, 0, byteRequestBody.Length);
                requestStream.Close();
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string contents = reader.ReadLine();

            }
            catch { }
        }

        /// <summary>
        ///  Method to close all browser associate with web driver 
        /// </summary>
        public void CloseAllBrowser()
        {
            try
            {
                 
                 if (Property.isSauceLabExecution.Equals(true))
                    {
                      try
                      {
                  
                        driver.Quit();
                        updateJobDeatils();
                      }
                     catch(Exception ec)
                      {
                      throw ec;
                      }
                    }
                    else
                    {
                        //GetWindowHandles();//this method set arrKnownBrowserHwnd in TestObject class. 
                        //IEnumerable<string> windowHandles = Property.arrKnownBrowserHwnd;
                        //foreach (string hwnd in windowHandles)
                        //{
                        //    try
                        //    {
                        //        driver.SwitchTo().Window(hwnd);
                        //        driver.Quit();
                        //    }
                        //    catch (Exception) { }
                        //}
                        driver.Quit();
                        System.Threading.Thread.Sleep(3000);
                    }               

                //Wait for profile to be free before moving on to next step
                if (browserName.Equals(KryptonConstants.BROWSER_FIREFOX))
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    Func<IWebDriver, bool> condition = this.ProfileRunningCondition;
                    try
                    {
                        wait.Until(condition);
                    }
                    catch
                    {
                       
                    }
                }
                
            }
          
            catch (WebDriverException e)
            {
               
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void setBrowserFocus()
        {
            try
            {
                Browser.switchToMostRecentBrowser();
                Browser.handleIECertificateError();
            }
            catch
            {
            }
        }

        public static string switchToMostRecentBrowser()
        {
            {

                GetWindowHandles();
                driver.SwitchTo().Window(Property.hwndMostRecentWindow);

                //If its ie and locally running, set focus as well
                if (browserName.ToLower().Equals("ie") && isRemoteExecution.ToString().ToLower().Equals("false"))
                {
                    activateCurrentBrowserWindow();
                }

            }

           

            //Bring current window to top of the page

            return Property.hwndMostRecentWindow;
        }


        public static string activateCurrentBrowserWindow()
        {
            // try to switch to most recent browser window, if you can
            try
            {
                string winTitle = driver.Title;
                AutoItX3Lib.AutoItX3 objAutoit = new AutoItX3Lib.AutoItX3();
                Object[,] windowList;
                windowList = (Object[,])objAutoit.WinList("[TITLE:" + winTitle + "; CLASS:IEFrame]");
                int windowCount = (int)windowList[0, 0];
                if (windowCount.Equals(1))
                {
                    //Setting window on top means you cannot manually switch to any other windows
                    //If you need IE clicks to be stable, this has to be done
                    objAutoit.WinSetOnTop(winTitle, "", 1);

                    //Activate window and set to focus
                    objAutoit.WinActivate(winTitle);
                }
            }
            catch (Exception)
            {
               
            }
            return Property.hwndMostRecentWindow;
        }


        public static void checkForApplicationCrash()
        {
            //Checking for application crash situations here.
            //In case of application crash, throw exception and exit test case execution
            try
            {
               

                string winTitleForCrashCheck = driver.Title;
                string htmlTextOnPage = driver.FindElement(By.TagName("h1")).Text;

                if (htmlTextOnPage.IndexOf("Server Error in",
                                                   StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Common.Property.ExecutionFailReason = "Application Crash: " + htmlTextOnPage;
                    Common.Property.EndExecutionFlag = true;
                    throw new Exception(Common.Property.ExecutionFailReason);
                }

            }
            catch (Exception crashCheck)
            {
                
            }
        }

        /// <summary>
        ///  Method to get all widow handles associated with current driver.
        /// </summary>
        public static void GetWindowHandles()
        {
            //Retrieve handles for all opened browser windows
            string recentWindowHandle = string.Empty;
            IEnumerable<string> hwndWindowsHandles = driver.WindowHandles.ToArray();

            if (hwndWindowsHandles.Count().Equals(1))
            {
                Property.arrKnownBrowserHwnd = hwndWindowsHandles;
                recentWindowHandle = Property.arrKnownBrowserHwnd.Last();
            }
            else
            {

                IEnumerable<string> arrNewWindos = hwndWindowsHandles.Except(Property.arrKnownBrowserHwnd);

                //when a new windows has appeared
                if (arrNewWindos.Count() >= 1)
                {
                    Property.arrKnownBrowserHwnd = Property.arrKnownBrowserHwnd.Union(hwndWindowsHandles);
                }

                //When no new windows appeared
                if (arrNewWindos.Count() == 0)
                {
                    Property.arrKnownBrowserHwnd = Property.arrKnownBrowserHwnd.Intersect(hwndWindowsHandles);
                }

                recentWindowHandle = Property.arrKnownBrowserHwnd.Last();
            }

            //Store window handles
            Common.Property.hwndFirstWindow = Property.arrKnownBrowserHwnd.First();
            Common.Property.hwndMostRecentWindow = Property.arrKnownBrowserHwnd.Last();

        }

        /// <summary>
        ///  Method to check whether firefoxprofile is still running or not.
        /// </summary>
        /// <param name="driver">IWebDriver</param>
        /// <returns>bool value</returns>
        private bool ProfileRunningCondition(IWebDriver driver)
        {
            try
            {
               
                if (!FirefoxDriverIsRunning()) // 20 Feb 2013
                    //This method is checking profile.
                    return true;
               
            }
            catch (Exception e)
            {

            }
            return true;
            
        }

        /// <summary>
        ///   Method to delete all cookie of associate web driver browser
        /// </summary>

        public void DeleteAllCookies()
        {
            try
            {
                
                try
                {
                    if(browser.isAlertPresent())
                    driver.SwitchTo().Alert().Accept(); 
                }
                catch { }
                if (signal==0)
               driver.Manage().Cookies.DeleteAllCookies();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
              
            }
        }

        /// <summary>
        ///   Method to open specified url in associate browser with web driver
        /// </summary>

        public void NavigationUrl(string url)
        {
            try
            {
                #region  Added for PBI 145
                bool IsSameInputUrlToNavigate = driver.Url.Equals(url);
                bool IsSameUrlAfterNavigate = false;
                string UrlBefoeNavigation = driver.Url;

                driver.Navigate().GoToUrl(url); // for chrome some time could not navigate to new url.

                IsSameUrlAfterNavigate = driver.Url.Equals(UrlBefoeNavigation);
                if (IsSameUrlAfterNavigate && !IsSameInputUrlToNavigate)
                {
                    if (!url.StartsWith(@"http://"))
                        url = @"http://" + url;
                    string navigateUrlScript = string.Format(" window.location='{0}'", url);
                    try
                    {
                        IJavaScriptExecutor chromejs = (IJavaScriptExecutor)driver;
                        chromejs.ExecuteScript(navigateUrlScript, null);
                    }
                    catch { }

                }

                #endregion

                DateTime startTime = DateTime.Now;

                //measure total time and raise exception if timeout is more than the allowed limit
                DateTime finishTime = DateTime.Now;
                double totalTime = (double)(finishTime - startTime).TotalSeconds;
                foreach (string modifiervalue in Common.Utility.driverKeydic.Values)
                {
                    if (modifiervalue.ToLower().Contains("timeout="))
                    {
                        double timeout = double.Parse(modifiervalue.Split('=').Last());
                        if (totalTime > timeout)
                        {
                            throw new Exception("Page load took " + totalTime.ToString() + " seconds to load against expected time of " + timeout + " seconds.");
                        }
                        else
                        {
                            Common.Property.Remarks = "Page load took " + totalTime.ToString() + " seconds to load against expected time of " + timeout + " seconds.";
                        }
                    }
                }

               
              
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                else if (e.Message.Contains(Common.exceptions.ERROR_404)) 
                {
                    throw new Exception("Page load is taking longer time.");
                }
                else
                    throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
                
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        /// <summary>
        ///  Method to navigate back in associate browser with web driver
        /// </summary>
        public void GoBack()
        {
            try
            {
                driver.Navigate().Back();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to navigate forward inassociate browser with web driver
        /// </summary>
        public void GoForward()
        {
            try
            {
                driver.Navigate().Forward();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        ///  Method to refersh the browser
        /// </summary>
        public void Refresh()
        {
            try
            {
                driver.Navigate().Refresh();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        ///  Method to  click on OK or Yes button in appeared dialog box.
        /// </summary>
        public void ClickDialogButton()
        {
            try
            {
                driver.SwitchTo().Alert().Accept();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        ///  Method to  click on OK or Yes button in appeared pop up.
        /// </summary>
        public void CloseBrowserPopUp()
        {
            try
            {

                driver.SwitchTo().Alert().Accept();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        ///  Method to accept condition specify in alert.
        /// </summary>
        public void AlertAccept()
        {
            try
            {
                //Implicitely Wait for specicified alert to appear.
                delAlertLoaded = waitAndGetAlert;
                alert = alertLoadingWait.Until(delAlertLoaded);
                alert.Accept();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                 throw new Exception("Could Not find Alert within " + Common.Property.Waitforalert + " seconds"); //Updated alert failure message 
            }
            catch (Exception)
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0011"));
            }
        }




        /// <summary>
        ///  Method to dismiss condition specify in alert
        /// </summary>
        public void AlertDismiss()
        {
            try
            {
                //Implicitely wait for an alert to appear.
                delAlertLoaded = waitAndGetAlert;
                alert = alertLoadingWait.Until(delAlertLoaded);
                alert.Dismiss();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0011"));
            }
        }
        /// <summary>
        /// This method will return true or false depending whether Text is present on alert or not
        /// </summary>
        /// <param name="text"></param>
        /// <param name="keyWordDic"></param>
        /// <returns>bool</returns>

        public bool VerifyAlertText(string text, Dictionary<int, string> keyWordDic)
        {
            bool result = false;
            try
            {
               
                    delAlertLoaded = waitAndGetAlert;

                    alert = alertLoadingWait.Until(delAlertLoaded);

                    string alertText = alert.Text;
                    if (!keyWordDic.Count.Equals(0))
                        result = Common.Utility.doKeywordMatch(text, alertText);
                    else
                        result = alertText.Equals(text);

                    //Return remarks if alert text matching failed. 
                    if (!result)
                    {
                        Common.Property.Remarks = "Actual alert text '" + alertText + "' does not matches with expected text '" + text + "'";
                    }

                    return result;

            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                  
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception("Could Not find Alert within " + Common.Property.Waitforalert+" seconds");
               
            }
            catch (Exception)
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0011"));
            }
        }

        /// <summary>
        /// Method to open new browser and initiate Driver with specified browser name.
        /// </summary>
        /// <param name="browserName">string : Name of the browser Ex."firefox"</param>
        /// <param name="deleteCookie">bool : This parameter indicate whether cookies should delete or not</param>
        /// <param name="url">string : Url of webpage to be open in new browser</param>
        /// <param name="isRemoteExecution">string : determine for remote execetion. value must be "true" or "false"</param>
        /// <param name="remoteUrl">string : Url of remote machine</param>
        /// <param name="browserDimension"> string: Dimension for Browser Window size.</param>>
        /// <returns></returns>
        public static Browser OpenBrowser(string browserName, bool deleteCookie, string url, 
            string isRemoteExecution, string remoteUrl, string ProfilePath, string addonsPath,
            DataSet datasetRecoverPopup, DataSet datasetRecoverBrowser, DataSet datasetOR, string browserDimension)
        {
            
            try
            {
                try
                {
                    if (!string.IsNullOrEmpty(browserDimension))
                    {
                        browserDimension = browserDimension.Trim().Split('=')[1];
                        width = int.Parse(browserDimension.Split(',')[0].Trim());
                        height = int.Parse(browserDimension.Split(',')[1].Trim());
                        IsBrowserDimendion = true;
                    }
                    else
                        IsBrowserDimendion = false;
                }
                catch
                {
                    throw new KryptonException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + "Input string was not in a correct format.");
                }
                //If cookies need to be deleted, no use of using a profile created by webdriver in previous instances
                if (deleteCookie)
                {
                    Common.Utility.SetVariable("FirefoxProfilePath", Common.Utility.GetParameter("FirefoxProfilePath"));
                    firefoxProfilePath = Common.Utility.GetParameter("FirefoxProfilePath");
                }
                if (Common.Property.parameterdic.ContainsKey(KryptonConstants.CHROME_PROFILE_PATH))
                {
                    if (!Path.IsPathRooted(Common.Property.parameterdic[KryptonConstants.CHROME_PROFILE_PATH]))
                    {
                        Common.Property.parameterdic[KryptonConstants.CHROME_PROFILE_PATH] = string.Concat(Common.Property.IniPath, Common.Property.parameterdic[KryptonConstants.CHROME_PROFILE_PATH]);
                        Common.Utility.SetVariable(KryptonConstants.CHROME_PROFILE_PATH, Common.Property.parameterdic[KryptonConstants.CHROME_PROFILE_PATH] = Path.GetFullPath(Common.Property.parameterdic[KryptonConstants.CHROME_PROFILE_PATH]));
                    }
                }
                                     
                //Determine which Driver will be initiated. 
                string prevBrowser = driverName;

                switch (browserName.ToLower())
                {
                    case KryptonConstants.BROWSER_FIREFOX:

                        Browser.firefoxProfilePath = ProfilePath;
                        driverName = KryptonConstants.FIREFOX_DRIVER;
                        break;
                    case KryptonConstants.BROWSER_IE:
                        driverName = KryptonConstants.INTERNET_EXPLORER_DRIVER;
                        break;
                    case KryptonConstants.BROWSER_CHROME:
                        Browser.chromeProfilePath = Common.Utility.GetVariable(KryptonConstants.CHROME_PROFILE_PATH);
                        driverName = KryptonConstants.CHROME_DRIVER;
                        break;
                    case KryptonConstants.BROWSER_SAFARI:
                        driverName = KryptonConstants.SAFARI_DRIVER;
                            break;
                }

                //Check url format.
                if (url.IndexOf(':') != 1 && !(url.Contains("http://") || url.Contains("https://"))) //Check for file protocol and http protocol :
                {
                    url = "http://" + url;
                }
                //If opening a file in Firefox.
                if (url.IndexOf(':').Equals(1) && url.IndexOf("xml", StringComparison.OrdinalIgnoreCase) >= 0 && !(url.Contains("http://") || url.Contains("https://")) && browserName.IndexOf("firefox", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    url = "View-source:" + url;
                }
                //Determine whether execution is remote or not.
                Browser.isRemoteExecution = isRemoteExecution;
                //Get Remote machine ip or name for remote execution.
                Browser.remoteUrl = remoteUrl;
                // Variable to determine whether to use firefox addons or not
                Browser.addonsPath = addonsPath;
                //Condition to check Whether to initiate new driver instance or just open new browser   

                int ExistingBrowserCount = 0;

                try
                {
                    if (driverName.ToLower().Equals(prevBrowser.ToLower()))
                    {
                        ExistingBrowserCount = driver.WindowHandles.Count;
                        driverlist.Add(driver);
                    }
                }
                catch { }
                
                if (browser == null || ExistingBrowserCount == 0)
                {
                    browser = new Browser(errorCaptureAs, browserName, deleteCookie);

                    browser.setBrowserFocus();                                    
                    browser.NavigationUrl(url);                 
                   if (deleteCookie && !browserName.ToLower().Equals(KryptonConstants.BROWSER_FIREFOX))
                    {
                        try
                        {
                            browser.DeleteAllCookies();
                        }
                            catch { }
                        browser.Refresh();
                        browser.NavigationUrl(url);
                    }
                }
                else
                {
                    //If there are browsers already open, launch with cookies instead
                    TestObject testObject = new TestObject();
                    testObject.ExecuteStatement("window.open();"); // Opening a browser using JavaScript              
                    browser.setBrowserFocus();
                    if (IsBrowserDimendion)
                      driver.Manage().Window.Size = new System.Drawing.Size(width, height);
                    
                    browser.NavigationUrl(url);

                    if (deleteCookie)
                    {
                       browser.DeleteAllCookies();
                       browser.NavigationUrl(url);
                    }
                }              
                return browser;

            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);

                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        ///  Method to wait for a specified time.
        /// </summary>
        /// <param name="time">string : Duration of time to wait</param>
        public void Wait(string time)
        {
            time = time.Trim();
            
            int waitTime = 0;
            try
            {
                if (string.IsNullOrEmpty(time))
                    time = "0";
                try
                {
                    waitTime = int.Parse(time);
                }
                catch (Exception e)
                {
                    waitTime = 0;
                }
                if (driver == null) // : Added to make wait function work with out opening the browser
                {

                    Thread.Sleep(1000 * (waitTime));

                }
                else
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitTime));
                    bool ExplicitWait = wait.Until<bool>((d) =>
                    {
                        return false;
                    });
                }
            }
            catch (WebDriverTimeoutException) { }
           
        }


        /// <summary>
        ///  Method to verify specied text on web page.
        /// </summary>
        /// <param name="text">Text to verify on web page</param>
        /// <returns>boolean value</returns>
        ///
        public bool VerifyTextPresentOnPage(string text, Dictionary<int, string> KeywordDic)
        {
            try
            {

                bool isKeyVerified = true;
                text = text.Trim();
                // Replace special characters here
                text = Common.Utility.ReplaceSpecialCharactersInString(text);
                try
                {
                    IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                    Browser.switchToMostRecentBrowser();
                }
                catch { }
                IWebElement element = driver.FindElement(By.XPath("//html"));
                string pageText = element.Text.ToString().Trim();

                if (!KeywordDic.Count.Equals(0))
                {
                    isKeyVerified = Common.Utility.doKeywordMatch(text, pageText);
                }
                else
                {
                    isKeyVerified = pageText.Contains(text);
                }
                if (!isKeyVerified)
                {
                    Common.Property.Remarks = "Text : \"" + text + "\" is not found";
                }
                return isKeyVerified;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e1)
            {           
                throw;
            }

        }

        /// <summary>
        ///  Method to return value of property if found other wise return null
        /// </summary>
        /// <param name="propertyType"> Type of property to get from test object.</param>
        /// <returns>Property value</returns>
        public string GetPageProperty(string propertyType)
        {
            try
            {
                try
                {
                    IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                    Browser.switchToMostRecentBrowser();
                }
                catch
                {
                // will not affect normal test case flow.
                }


                switch (propertyType.ToLower())
                {
                    case "title":
                        return driver.Title;
                    case "url":
                        return driver.Url;
                    default:
                        return null; ;
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        ///  Method to verify page property and return true if property found else return false
        /// </summary>
        /// <param name="pagePropertyType">Type of property to verify</param>
        /// <param name="pageProperty">Value of property to verify</param>
        /// <returns>boolean </returns>
        public bool VerifyPageProperty(string property, string propertyValue, Dictionary<int, string> keywordDic)
        {
            try
            {              
                string actualValue;
                actualValue = this.GetPageProperty(property);               
                bool status;
                if (!keywordDic.Count.Equals(0))
                {
                   
                    status = Common.Utility.doKeywordMatch(propertyValue, actualValue);                   
                    if (!status)
                    {

                        Common.Property.Remarks = "Keyword Match failed.Actual page property - \"" + actualValue + "\" does not match with expected page property - \"" + propertyValue + "\".";
                    }
                }
                else
                {
                    
                    status = actualValue.Equals(propertyValue);                
                       
                    
                }
                if (!status)
                {

                    Common.Property.Remarks = "Actual page property - \"" + actualValue + "\" does not match with expected page property - \"" + propertyValue + "\".";
                }
                return status;
            }

            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }

        }




        /// <summary>
        ///  Method to Verify specified text is present in web page view source code .
        /// </summary>
        /// <param name="text">Text to verify</param>
        /// <returns>boolean value</returns>
        public bool VerifyTextInPageSource(string text, Dictionary<int, string> KeyWordDic)
        {
            try
            {
                try
                {
                    IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                    Browser.switchToMostRecentBrowser();
                }
                catch { }

                bool isKeyVerified;
                string s = driver.PageSource;
                s = s.Replace("<br />", "<br>");
                s = s.Replace("<BR/>", "<br>");
                if (!KeyWordDic.Count.Equals(0))
                {
                    isKeyVerified = Common.Utility.doKeywordMatch(text, s);
                }
                else
                {
                    isKeyVerified = s.Contains(text);
                }
                if (!isKeyVerified)
                {
                    Common.Property.Remarks = "Actual text : \"" + text + "\" is not found";
                }
                return isKeyVerified;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw;
            }

        }


        /// <summary>
        ///  Method to Verify specified text is not present in web page view source code .
        /// </summary>
        /// <param name="text">text to verify</param>
        /// <returns>boolean value</returns>
        public bool VerifyTextNotOnPageSource(string text, Dictionary<int, string> keyWordDic)
        {
            try
            {
                bool status = this.VerifyTextInPageSource(text, keyWordDic);
                Common.Property.Remarks = string.Empty;
                if (status)
                {
                    Common.Property.Remarks = "Text : \"" + text + "\" is found.";
                }
                return (!status);
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        ///  Method to Method to verify page display by verifying page title.
        /// </summary>
        /// <param name="title">Title of web page</param>
        /// <returns>boolean value</returns>
        public bool VerifyPageDisplayed(Dictionary<int, string> keyWorddic)
        {

            try
            {
                bool isKeyVerified;
                string url = this.GetData(KryptonConstants.WHAT);
                if (!keyWorddic.Count.Equals(0))
                {
                    try
                    {
                        IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                        wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                        Browser.switchToMostRecentBrowser();
                    }
                    catch { }
                    isKeyVerified = Common.Utility.doKeywordMatch(url, driver.Url);
                }
                else
                {
                    //By Default regular expression match would be done for Url entry in OR.
                    try
                    {
                        IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                        wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                        Browser.switchToMostRecentBrowser();
                    }
                    catch { }
                    try
                    {
                        driver.SwitchTo().DefaultContent();
                    }
                    catch (Exception ed)
                    {
                        //do nothing
                    }
                    Regex regExp = new Regex(url.ToLower());
                    Match m = regExp.Match((driver.Url).ToLower());
                    isKeyVerified = m.Success;
                }

                if (!isKeyVerified)
                {
                    Common.Property.Remarks = "Page with URL: \"" + url + "\" is not displayed." +
                                              "Actual displayed page was: " + driver.Url;
                }

                return isKeyVerified;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public Boolean isAlertPresent()
        {

            Boolean presentFlag = false;

            try
            {

                // Check the presence of alert
                if (driver != null)
                {
                    IAlert alert = driver.SwitchTo().Alert();
                    // Alert present; set the flag
                    presentFlag = true;
                    
                }

            }
            catch (Exception ex)
            {
                string s;
                s = ex.StackTrace;
            }

            return presentFlag;

        }
        /// <summary>
        /// Method to Method to create printscreen and ViewSource file of webpage.
        /// </summary>
        /// <param name="stepActionNumber">Step number of action to be executed</param>
        /// <param name="path">Path where image and html source file will be save</param>
        /// <returns></returns>
        public string GetScreenShot(string stepActionNumber, string path, string stepAction = "")
        {
            //if alert exists then will not be taken screen shot
            if (isAlertPresent())
            {
                return string.Empty;
            }
            if (driver == null)
                return string.Empty;
            Screenshot screenShot = null;
            string imageName = string.Empty;
            string htmlFileName = string.Empty;
            try
            {
                if ((errorCaptureAs.ToLower().Equals("both") || errorCaptureAs.ToLower().Equals("image")))
                {
                   
                    switchToMostRecentBrowser();

                    //Get current url of the page
                    Property.AttachmentsUrl = driver.Url;

                    if (isRemoteExecution.Equals("true"))
                    {
                        try
                        {
                            screenShot = ((ITakesScreenshot)driver).GetScreenshot();
                            
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Take snapshot: " + e.Message);
                        }
                    }
                    else
                    {
                        switch (browserName.ToLower())
                        {
                            case KryptonConstants.BROWSER_FIREFOX:
                                FirefoxDriver firefoxDriver = (FirefoxDriver)driver;
                                screenShot = ((ITakesScreenshot)firefoxDriver).GetScreenshot();
                                break;
                            case KryptonConstants.BROWSER_IE:
                                InternetExplorerDriver ieDriver = (InternetExplorerDriver)driver;
                                try
                                {
                                    screenShot = ((ITakesScreenshot)ieDriver).GetScreenshot();
                                }
                                catch
                                {
                                    return " | ";
                                }
                                break;
                            case KryptonConstants.BROWSER_CHROME:
                                //Updated Screenshot capturing in case of ChromeDriver.
                                ChromeDriver chromeDriver = (ChromeDriver)driver;
                                screenShot = ((ITakesScreenshot)chromeDriver).GetScreenshot();
                                break;
                            default:
                                return "Invalid browser string";

                        }
                    }

                    imageName = "Image" + stepActionNumber + ".jpg";
                    string imageNameHQ = "Image.jpg"; // High Quality Image

                    //Create Failed step image.
                    if (screenShot != null)
                    {
                        string screenshot = screenShot.AsBase64EncodedString;
                        byte[] screenshotAsByteArray = screenShot.AsByteArray;

                        //Saving HQ image first
                        screenShot.SaveAsFile(path + "/" + imageNameHQ, System.Drawing.Imaging.ImageFormat.Jpeg);
                        // Compress Image
                        try
                        {
                            Common.Utility.CompressImage(path + "/" + imageNameHQ, path + "/" + imageName);
                            // Delete HQ image after compression
                            File.Delete(path + "/" + imageNameHQ);
                        }
                        catch
                        {

                        }
                    }
                }
                if (errorCaptureAs.ToLower().Equals("both") || errorCaptureAs.ToLower().Equals("html"))
                {
                    try
                    {
                        string pageViewSource = driver.PageSource;
                        htmlFileName = "Html" + stepActionNumber + ".html";
                        StreamWriter sw = new StreamWriter(path + "/" + htmlFileName, false, Encoding.UTF8);
                        //Create html view source file.
                        sw.WriteLine(pageViewSource.ToCharArray());
                        sw.Close();
                    }
                    catch
                    {

                    }
                }
                return imageName + "|" + htmlFileName;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains("No response from server for url"))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        /// <summary>
        ///Method to Quit initialized web driver
        /// </summary>
        public void Shutdown()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(driver.Title) && Property.isSauceLabExecution.Equals(true))
                {
                    //Saucelabs assume driver as diconnected only when using quit().
                    try { driver.Quit(); }
                    catch { }
                    driver.Close();
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Common.exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        ///  Method to get the instance of selenium version 1.3
        /// </summary>
        /// <returns>Selenium instance</returns>
        public static Selenium.DefaultSelenium GetSeleniumOne()
        {
            try
            {
                if (selenium == null)
                {
                    selenium = new Selenium.WebDriverBackedSelenium(driver, driver.Url);
                }
                try
                {
                    selenium.Start();
                }
                catch (Exception)
                {

                }
                return selenium;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void executeSauceConnect()
        {

            try
            {

                foreach (Process proc in Process.GetProcessesByName("sc"))
                {
                    proc.Kill();
                }
            }

            catch
            { }


            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.FileName = Common.Utility.GetParameter("SauceConnectPath");
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = "-u " + Common.Utility.GetParameter("username") + " " + "-k " + Common.Utility.GetParameter("password") + Common.Utility.GetParameter("CommandLineOptions");

                try
                {
                    using (Process exeProcess = Process.Start(startInfo)) ;
                }

                catch
                { }

            }

            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }

    }
}
