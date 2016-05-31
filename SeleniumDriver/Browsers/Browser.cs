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
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Common;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace Driver.Browsers
{

    public class Browser
    {

        public static IWebDriver Driver;
        public static string BrowserName;
        //public static string browserVersion;
        private Dictionary<string, string> _objDataRow = new Dictionary<string, string>();
        private static Selenium.DefaultSelenium _selenium;
        private static string _errorCaptureAs = string.Empty;
        public static string IsRemoteExecution = string.Empty;
        private static string _remoteUrl = string.Empty;
        private static string _addonsPath = string.Empty;    // Addons path is get from parameters.ini
        private IAlert _alert;
        public WebDriverWait AlertLoadingWait; //used for asynchronous wait for alerts.
        public Func<IWebDriver, IAlert> DelAlertLoaded;
        public string WaitTimeForAlerts = string.Empty;
        private static Browser _browser;
        private static string _driverName = string.Empty;
        public static List<IWebDriver> Driverlist = new List<IWebDriver>();  // to maintain all opened browser(used in close Browser method)
        public static int Width = 600; // default values for Browser Window
        public static int Height = 800;
        SauceLabs _objSauceLabs = new SauceLabs();
        public static Boolean IsBrowserDimendion;
        //Advanced user Interaction API object for driver
        public static Actions DriverActions;
        public static int Signal = 0;


        /// <summary>
        ///  Constructor to call initDriver method and set browser name variable.
        /// </summary>
        public Browser(string errorCapture, string browserName = null, bool deleteCookie = true)
        {

            if (string.IsNullOrWhiteSpace(browserName) == false)
            {
                BrowserName = browserName;
                InitDriver(deleteCookie);
            }

            if (string.IsNullOrWhiteSpace(errorCapture) == false)
            {
                _errorCaptureAs = errorCapture;
            }
            else
            {
                errorCapture = "image";
            }
        }

        /// <summary>
        ///  set object information dictionary.
        /// </summary>
        /// <param name="objDataRow">Dictionary : Test Object information dictionary</param>
        public void SetObjDataRow(Dictionary<string, string> objDataRow)
        {
            _objDataRow = objDataRow;
            // Refrencing waitLoading Object from TestObject.
            AlertLoadingWait = TestObject.ObjloadingWait;
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
                        return _objDataRow["logical_name"];
                    case KryptonConstants.OBJ_TYPE:
                        return _objDataRow[KryptonConstants.OBJ_TYPE];
                    case KryptonConstants.HOW:
                        return _objDataRow[KryptonConstants.HOW];
                    case KryptonConstants.WHAT:
                        return _objDataRow[KryptonConstants.WHAT];
                    case KryptonConstants.MAPPING:
                        return _objDataRow[KryptonConstants.MAPPING];
                    default:
                        return null;
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (KeyNotFoundException)
            {
                throw new KeyNotFoundException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0008"));
            }
        }

        /// <summary>
        /// Assigning platform during execution.
        /// </summary>
        /// <returns>Platform Type Object.</returns>
        private PlatformType AssignPlatform()
        {
            PlatformType platform;
            var platformString = Utility.GetParameter("Platform");
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
        private void InitDriver(bool deleteCookie = true)
        {
            try
            {
                Property.IsSauceLabExecution = false;
                if (String.Compare(IsRemoteExecution, "false", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    switch (BrowserName.ToLower())
                    {
                        case KryptonConstants.BROWSER_FIREFOX:
                            new FireFox().IntiailizeDriver(ref Driver, ref IsBrowserDimendion, ref Driverlist, ref Width, ref Height, ref _addonsPath);
                            break;
                        case KryptonConstants.BROWSER_IE:
                            new IE().IntializeDriver(ref Driver, ref IsBrowserDimendion, ref Driverlist, ref Width, ref Height, ref deleteCookie);
                            break;
                        case KryptonConstants.BROWSER_CHROME:
                             new Chrome().IntializeDriver(ref  Driver, ref  IsBrowserDimendion, ref  Driverlist, ref  Width, ref  Height);
                            break;
                        case KryptonConstants.BROWSER_SAFARI:
                            new Safari().IntializeDriver(ref Driver, ref IsBrowserDimendion, ref Driverlist, ref Width, ref Height);
                            break;
                        default:
                            Console.WriteLine("No browser is defined.");
                            break;
                    }
                }
                else
                {
                    new SauceLabs().IntializeDriver(ref _remoteUrl, ref BrowserName, ref Driver, ref DriverActions);
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL) || e.Message.IndexOf("Connection refused", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (IsRemoteExecution.Equals("true"))
                    {
                        InitDriver();
                    }
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw;
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf(Exceptions.ERROR_PARSINGVALUE, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (IsRemoteExecution.Equals("true"))
                    {
                        InitDriver();
                    }
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw;
            }
        }


        /// <summary>
        /// WaitAndGetAlert will wait until expected alert is returned OR specified time has been exceeded.
        /// </summary>
        /// <param name="ldriver">IwebDriver instance</param>
        /// <returns>IAlert : Alert Interface </returns>
        public IAlert WaitAndGetAlert(IWebDriver ldriver)
        {

            double totalSeconds = (double)DateTime.Now.Ticks / TimeSpan.TicksPerSecond;

            while (((double)DateTime.Now.Ticks / TimeSpan.TicksPerSecond - totalSeconds) <= Property.Waitforalert)
            {
                if (IsAlertPresent())
                {
                    return Driver.SwitchTo().Alert();
                }
            }
            return null;
        }



        public static string activateCurrentBrowserWindow()
        {
            // try to switch to most recent browser window, if you can
            try
            {
                string winTitle = Driver.Title;
                AutoItX3Lib.AutoItX3 objAutoit = new AutoItX3Lib.AutoItX3();
                var windowList = (Object[,])objAutoit.WinList("[TITLE:" + winTitle + "; CLASS:IEFrame]");
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
                // Ignore
            }
            return Property.HwndMostRecentWindow;
        }


        public static string switchToMostRecentBrowser()
        {
            {

                GetWindowHandles();
                Driver.SwitchTo().Window(Property.HwndMostRecentWindow);

                //If its ie and locally running, set focus as well
                if (BrowserName.ToLower().Equals("ie") && IsRemoteExecution.ToLower().Equals("false"))
                {
                    activateCurrentBrowserWindow();
                }

            }
            //Bring current window to top of the page

            return Property.HwndMostRecentWindow;
        }


        /// <summary>
        /// Method to close current browser associate with web driver 
        /// </summary>
        public void CloseBrowser()
        {
            try
            {
                try
                {
                    switchToMostRecentBrowser();
                }
                catch
                {
                    // ignored
                }
                try
                {

                    Driver.Close();
                    Thread.Sleep(3000);
                    if (Driverlist.Count > 1)
                    {
                        for (int i = Driverlist.Count - 1; i >= 0; i--)
                        {
                            try
                            {
                                Driver = Driverlist[i];
                                string window = switchToMostRecentBrowser();
                                Driver = Driver.SwitchTo().Window(window);
                                Driver.Title.ToString();  // Test for Exception with current driver.
                                break;
                            }
                            catch (Exception)
                            {
                                // ignored
                            }
                        }
                        if (Driver.ToString().IndexOf("InternetExplorer", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            BrowserName = "ie";
                            _driverName = "InternetExplorerDriver";

                        }
                        else if (Driver.ToString().IndexOf("chrome", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _driverName = "ChromeDriver";
                            BrowserName = "chrome";
                        }
                        else if (Driver.ToString().IndexOf("Firefox", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _driverName = "FirefoxDriver";
                            BrowserName = "Firefox";
                        }
                        else if (Driver.ToString().IndexOf("Safari", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            _driverName = "SafariDriver";
                            BrowserName = "Safari";
                        }
                        TestObject.Driver = Driver;
                        Driver = TestObject.Driver;
                    }
                }
                catch (Exception)
                {
                    //Do nothing No browser to close.
                }
                if (BrowserName.ToLower().Equals("firefox") && Property.ArrKnownBrowserHwnd.Count().Equals(1))
                {
                    WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                    Func<IWebDriver, bool> condition = ProfileRunningCondition;
                    try
                    {
                        wait.Until(condition);
                    }
                    catch
                    {
                        // ignored
                    }
                }

            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains("No response from server for url"))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }

            catch (NullReferenceException e)
            {
                throw new NullReferenceException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0010") + ":" + e.Message);
            }

            catch (Exception e)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }




        /// <summary>
        ///  Method to check running firefox profile condition.
        /// </summary>
        private void FreeFireFoxProfile()
        {
            FireFox objFireFox = new FireFox();
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
            Func<IWebDriver, bool> condition = ProfileRunningCondition;
            wait.Until(condition);
        }

        /// <summary>
        ///  Method to check whether firefoxprofile is still running or not.
        /// </summary>
        /// <param name="driver">IWebDriver</param>
        /// <returns>bool value</returns>
        internal bool ProfileRunningCondition(IWebDriver driver)
        {
            if (!FirefoxDriverIsRunning())
                //This method is checking profile.
                return true;
            return false;

        }


        private Boolean FirefoxDriverIsRunning()
        {
            return Process.GetProcesses().Any(proc => proc.ProcessName.ToString().ToLower() == "firefox" || proc.ProcessName.ToString().ToLower() == "firefoxdriver");
        }


        /// <summary>
        ///  Method to close all browser associate with web driver 
        /// </summary>
        public void CloseAllBrowser()
        {
            try
            {

                if (Property.IsSauceLabExecution.Equals(true))
                {
                    Driver.Quit();
                    _objSauceLabs.UpdateJobDeatils();
                }
                else
                {
                    GetWindowHandles();//this method set arrKnownBrowserHwnd in TestObject class. 
                    IEnumerable<string> windowHandles = Property.ArrKnownBrowserHwnd;
                    foreach (string hwnd in windowHandles)
                    {

                        try
                        {
                            Driver.SwitchTo().Window(hwnd);
                            Driver.Quit();
                        }
                        catch (Exception)
                        {
                            // ignored
                        }
                    }
                }

                //Wait for profile to be free before moving on to next step
                if (BrowserName.Equals("firefox"))
                {
                    WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                    Func<IWebDriver, bool> condition = ProfileRunningCondition;
                    try
                    {
                        wait.Until(condition);
                    }
                    catch
                    {
                        // ignored
                    }
                }

            }

            catch (WebDriverException e)
            {

            }
        }

        public void SetBrowserFocus()
        {
            SwitchToMostRecentBrowser();
            IE.handleIECertificateError(ref Driver);
        }

        public static string SwitchToMostRecentBrowser()
        {
            try
            {
                GetWindowHandles();
                Driver.SwitchTo().Window(Property.HwndMostRecentWindow);
                //If its ie and locally running, set focus as well
                if (BrowserName.ToLower().Equals("ie") && IsRemoteExecution.ToString().ToLower().Equals("false"))
                {
                    ActivateCurrentBrowserWindow();
                }
                return Property.HwndMostRecentWindow;
            }
            catch (WebDriverException e)
            {
                KryptonException.Writeexception(e);
            }
            return Property.HwndMostRecentWindow;
        }


        public static string ActivateCurrentBrowserWindow()
        {
            // try to switch to most recent browser window, if you can
            try
            {
                string winTitle = Driver.Title;
                AutoItX3Lib.AutoItX3 objAutoit = new AutoItX3Lib.AutoItX3();
                var windowList = (Object[,])objAutoit.WinList("[TITLE:" + winTitle + "; CLASS:IEFrame]");
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
            catch (Exception ex)
            {
                KryptonException.Writeexception(ex);
            }
            return Property.HwndMostRecentWindow;
        }


        public static void CheckForApplicationCrash()
        {
            //Checking for application crash situations here.
            //In case of application crash, throw exception and exit test case execution
            try
            {
                string winTitleForCrashCheck = Driver.Title;
                string htmlTextOnPage = Driver.FindElement(By.TagName("h1")).Text;

                if (htmlTextOnPage.IndexOf("Server Error in",
                                                   StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Property.ExecutionFailReason = "Application Crash: " + htmlTextOnPage;
                    Property.EndExecutionFlag = true;
                    throw new Exception(Property.ExecutionFailReason);
                }

            }
            catch (Exception crashCheck)
            {
                KryptonException.Writeexception(crashCheck);
            }
        }

        /// <summary>
        ///  Method to get all widow handles associated with current driver.
        /// </summary>
        public static void GetWindowHandles()
        {
            //Retrieve handles for all opened browser windows
            IEnumerable<string> hwndWindowsHandles = Driver.WindowHandles.ToArray();

            if (hwndWindowsHandles.Count().Equals(1))
            {
                Property.ArrKnownBrowserHwnd = hwndWindowsHandles;
            }
            else
            {
                IEnumerable<string> arrNewWindos = hwndWindowsHandles.Except(Property.ArrKnownBrowserHwnd);
                //when a new windows has appeared
                if (arrNewWindos.Count() >= 1)
                {
                    Property.ArrKnownBrowserHwnd = Property.ArrKnownBrowserHwnd.Union(hwndWindowsHandles);
                }

                //When no new windows appeared
                if (arrNewWindos.Count() == 0)
                {
                    Property.ArrKnownBrowserHwnd = Property.ArrKnownBrowserHwnd.Intersect(hwndWindowsHandles);
                }
            }

            //Store window handles
            Property.HwndFirstWindow = Property.ArrKnownBrowserHwnd.First();
            Property.HwndMostRecentWindow = Property.ArrKnownBrowserHwnd.Last();
        }
        
        /// <summary>
        ///   Method to delete all cookie of associate web driver browser
        /// </summary>
        public void DeleteAllCookies()
        {
            try
            {
                if(_browser.IsAlertPresent()) 
                    Driver.SwitchTo().Alert().Accept();
                if (Signal == 0)
                    Driver.Manage().Cookies.DeleteAllCookies();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                KryptonException.Writeexception(e);
            }
        }

        /// <summary>
        ///   Method to open specified url in associate browser with web driver
        /// </summary>
        public void NavigationUrl(string url)
        {
            try
            {
                bool isSameInputUrlToNavigate = Driver.Url.Equals(url);
                bool isSameUrlAfterNavigate = false;
                string urlBefoeNavigation = Driver.Url;

                Driver.Navigate().GoToUrl(url); // for chrome some time could not navigate to new url.

                isSameUrlAfterNavigate = Driver.Url.Equals(urlBefoeNavigation);
                if (isSameUrlAfterNavigate && !isSameInputUrlToNavigate)
                {
                    if (!url.StartsWith(@"http://"))
                        url = @"http://" + url;
                    string navigateUrlScript = string.Format(" window.location='{0}'", url);
                    IJavaScriptExecutor chromejs = (IJavaScriptExecutor)Driver;
                    chromejs.ExecuteScript(navigateUrlScript, null);
                }
                DateTime startTime = DateTime.Now;
                //measure total time and raise exception if timeout is more than the allowed limit
                DateTime finishTime = DateTime.Now;
                double totalTime = (finishTime - startTime).TotalSeconds;
                foreach (string modifiervalue in Utility.DriverKeydic.Values)
                {
                    if (modifiervalue.ToLower().Contains("timeout="))
                    {
                        double timeout = double.Parse(modifiervalue.Split('=').Last());
                        if (totalTime > timeout)
                        {
                            throw new Exception("Page load took " + totalTime + " seconds to load against expected time of " + timeout + " seconds.");
                        }
                        Property.Remarks = "Page load took " + totalTime + " seconds to load against expected time of " + timeout + " seconds.";
                    }
                }



            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                if (e.Message.Contains(Exceptions.ERROR_404))
                {
                    throw new Exception("Page load is taking longer time.");
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }

        /// <summary>
        ///  Method to navigate back in associate browser with web driver
        /// </summary>
        public void GoBack()
        {
            try
            {
                Driver.Navigate().Back();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }

        /// <summary>
        /// Method to navigate forward inassociate browser with web driver
        /// </summary>
        public void GoForward()
        {
            try
            {
                Driver.Navigate().Forward();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }


        /// <summary>
        ///  Method to refersh the browser
        /// </summary>
        public void Refresh()
        {
            try
            {
                Driver.Navigate().Refresh();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        ///  Method to  click on OK or Yes button in appeared dialog box.
        /// </summary>
        public void ClickDialogButton()
        {
            try
            {
                Driver.SwitchTo().Alert().Accept();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }


        /// <summary>
        ///  Method to  click on OK or Yes button in appeared pop up.
        /// </summary>
        public void CloseBrowserPopUp()
        {
            try
            {
                Driver.SwitchTo().Alert().Accept();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
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
                DelAlertLoaded = WaitAndGetAlert;
                _alert = AlertLoadingWait.Until(DelAlertLoaded);
                _alert.Accept();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception("Could Not find Alert within " + Property.Waitforalert + " seconds"); //Updated alert failure message 
            }
            catch (Exception)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0011"));
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
                DelAlertLoaded = WaitAndGetAlert;
                _alert = AlertLoadingWait.Until(DelAlertLoaded);
                _alert.Dismiss();
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0011"));
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
            try
            {

                DelAlertLoaded = WaitAndGetAlert;

                _alert = AlertLoadingWait.Until(DelAlertLoaded);

                string alertText = _alert.Text;
                var result = false;
                result = !keyWordDic.Count.Equals(0) ? Utility.DoKeywordMatch(text, alertText) : alertText.Equals(text);

                //Return remarks if alert text matching failed. 
                if (!result)
                {
                    Property.Remarks = "Actual alert text '" + alertText + "' does not matches with expected text '" + text + "'";
                }

                return result;

            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {

                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception("Could Not find Alert within " + Property.Waitforalert + " seconds");

            }
            catch (Exception)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0011"));
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
        /// <param name="browserDimension"> string: Dimension for Browser Window size.</param>
        /// <param name="profilePath">Path to the browser profile if available.</param>
        /// <param name="addonsPath"></param>
        /// >
        /// <returns></returns>
        public static Browser OpenBrowser(string browserName, bool deleteCookie, string url,
            string isRemoteExecution, string remoteUrl, string profilePath, string addonsPath,
            DataSet datasetRecoverPopup, DataSet datasetRecoverBrowser, DataSet datasetOR, string browserDimension)
        {
            try
            {
                try
                {
                    if (!string.IsNullOrEmpty(browserDimension))
                    {
                        browserDimension = browserDimension.Trim().Split('=')[1];
                        Width = int.Parse(browserDimension.Split(',')[0].Trim());
                        Height = int.Parse(browserDimension.Split(',')[1].Trim());
                        IsBrowserDimendion = true;
                    }
                    else
                        IsBrowserDimendion = false;
                }
                catch
                {
                    throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + "Input string was not in a correct format.");
                }
                //If cookies need to be deleted, no use of using a profile created by webdriver in previous instances
                if (deleteCookie)
                {
                    Utility.SetVariable("FirefoxProfilePath", Utility.GetParameter("FirefoxProfilePath"));
                    FireFox.FirefoxProfilePath = Utility.GetParameter("FirefoxProfilePath");
                }
                if (Property.Parameterdic.ContainsKey(KryptonConstants.CHROME_PROFILE_PATH))
                {
                    if (!Path.IsPathRooted(Property.Parameterdic[KryptonConstants.CHROME_PROFILE_PATH]))
                    {
                        Property.Parameterdic[KryptonConstants.CHROME_PROFILE_PATH] = string.Concat(Property.IniPath, Property.Parameterdic[KryptonConstants.CHROME_PROFILE_PATH]);
                        Utility.SetVariable(KryptonConstants.CHROME_PROFILE_PATH, Property.Parameterdic[KryptonConstants.CHROME_PROFILE_PATH] = Path.GetFullPath(Property.Parameterdic[KryptonConstants.CHROME_PROFILE_PATH]));
                    }
                }

                //Determine which Driver will be initiated. 
                string prevBrowser = _driverName;
                switch (browserName.ToLower())
                {
                    case KryptonConstants.BROWSER_FIREFOX:

                        global::Driver.Browsers.FireFox.FirefoxProfilePath = profilePath;
                        _driverName = KryptonConstants.FIREFOX_DRIVER;
                        break;
                    case KryptonConstants.BROWSER_IE:
                        _driverName = KryptonConstants.INTERNET_EXPLORER_DRIVER;
                        break;
                    case KryptonConstants.BROWSER_CHROME:
                        Chrome.ChromeProfilePath = Utility.GetVariable(KryptonConstants.CHROME_PROFILE_PATH);
                        _driverName = KryptonConstants.CHROME_DRIVER;
                        break;
                    case KryptonConstants.BROWSER_SAFARI:
                        _driverName = KryptonConstants.SAFARI_DRIVER;
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
                IsRemoteExecution = isRemoteExecution;
                //Get Remote machine ip or name for remote execution.
                _remoteUrl = remoteUrl;
                //Variable to determine whether to use firefox addons or not
                _addonsPath = addonsPath;
                //Condition to check Whether to initiate new driver instance or just open new browser.   
                int existingBrowserCount = 0;
                try
                {
                    if (_driverName.ToLower().Equals(prevBrowser.ToLower()))
                    {
                        existingBrowserCount = Driver.WindowHandles.Count;
                        Driverlist.Add(Driver);
                    }
                }
                catch (Exception e)
                {
                    KryptonException.Writeexception(e);
                }
                
                if (_browser == null || existingBrowserCount == 0)
                {
                    _browser = new Browser(_errorCaptureAs, browserName, deleteCookie);

                    _browser.SetBrowserFocus();
                    _browser.NavigationUrl(url);
                    if (deleteCookie && !browserName.ToLower().Equals(KryptonConstants.BROWSER_FIREFOX))
                    {
                        _browser.DeleteAllCookies();
                        _browser.Refresh();
                        _browser.NavigationUrl(url);
                    }
                }
                else
                {
                    //If there are browsers already open, launch with cookies instead
                    TestObject testObject = new TestObject();
                    testObject.ExecuteStatement("window.open();"); // Opening a browser using JavaScript              
                    _browser.SetBrowserFocus();
                    if (IsBrowserDimendion)
                        Driver.Manage().Window.Size = new System.Drawing.Size(Width, Height);

                    _browser.NavigationUrl(url);

                    if (deleteCookie)
                    {
                        _browser.DeleteAllCookies();
                        _browser.NavigationUrl(url);
                    }
                }
                return _browser;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);

                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
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

            try
            {
                if (string.IsNullOrEmpty(time))
                    time = "0";
                int waitTime;
                try
                {
                    waitTime = int.Parse(time);
                }
                catch (Exception)
                {
                    waitTime = 0;
                }
                if (Driver == null) //Added to make wait function work with out opening the browser
                {

                    Thread.Sleep(1000 * (waitTime));

                }
                else
                {
                    WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(waitTime));
                    bool explicitWait = wait.Until<bool>((d) => false);
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
                text = Utility.ReplaceSpecialCharactersInString(text);
                try
                {
                    IWait<IWebDriver> wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete"));
                    SwitchToMostRecentBrowser();
                }
                catch
                {
                    // ignored
                }
                IWebElement element = Driver.FindElement(By.XPath("//html"));
                string pageText = element.Text.Trim();

                if (!KeywordDic.Count.Equals(0))
                {
                    isKeyVerified = Utility.DoKeywordMatch(text, pageText);
                }
                else
                {
                    isKeyVerified = pageText.Contains(text);
                }
                if (!isKeyVerified)
                {
                    Property.Remarks = "Text : \"" + text + "\" is not found";
                }
                return isKeyVerified;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
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
                    IWait<IWebDriver> wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete"));
                    SwitchToMostRecentBrowser();
                }
                catch
                {
                    // will not affect normal test case flow.
                }


                switch (propertyType.ToLower())
                {
                    case "title":
                        return Driver.Title;
                    case "url":
                        return Driver.Url;
                    default:
                        return null; ;
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }

        /// <summary>
        ///  Method to verify page property and return true if property found else return false
        /// </summary>
        /// <returns>boolean </returns>
        public bool VerifyPageProperty(string property, string propertyValue, Dictionary<int, string> keywordDic)
        {
            try
            {
                var actualValue = GetPageProperty(property);
                bool status;
                if (!keywordDic.Count.Equals(0))
                {

                    status = Utility.DoKeywordMatch(propertyValue, actualValue);
                    if (!status)
                    {

                        Property.Remarks = "Keyword Match failed.Actual page property - \"" + actualValue + "\" does not match with expected page property - \"" + propertyValue + "\".";
                    }
                }
                else
                {

                    status = actualValue.Equals(propertyValue);


                }
                if (!status)
                {

                    Property.Remarks = "Actual page property - \"" + actualValue + "\" does not match with expected page property - \"" + propertyValue + "\".";
                }
                return status;
            }

            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
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
        /// <param name="keyWordDic"></param>
        /// <returns>boolean value</returns>
        public bool VerifyTextInPageSource(string text, Dictionary<int, string> keyWordDic)
        {
            try
            {
                try
                {
                    IWait<IWebDriver> wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30.00));
                    wait.Until(driver1 => ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete"));
                    SwitchToMostRecentBrowser();
                }
                catch { }

                bool isKeyVerified;
                string s = Driver.PageSource;
                s = s.Replace("<br />", "<br>");
                s = s.Replace("<BR/>", "<br>");
                if (!keyWordDic.Count.Equals(0))
                {
                    isKeyVerified = Utility.DoKeywordMatch(text, s);
                }
                else
                {
                    isKeyVerified = s.Contains(text);
                }
                if (!isKeyVerified)
                {
                    Property.Remarks = "Actual text : \"" + text + "\" is not found";
                }
                return isKeyVerified;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }

        /// <summary>
        ///  Method to Verify specified text is not present in web page view source code .
        /// </summary>
        /// <param name="text">text to verify</param>
        /// <param name="keyWordDic"></param>
        /// <returns>boolean value</returns>
        public bool VerifyTextNotOnPageSource(string text, Dictionary<int, string> keyWordDic)
        {
            bool status = VerifyTextInPageSource(text, keyWordDic);
            Property.Remarks = string.Empty;
            if (status)
            {
                Property.Remarks = "Text : \"" + text + "\" is found.";
            }
            return (!status);
        }

        /// <summary>
        ///  Method to Method to verify page display by verifying page title.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyPageDisplayed(Dictionary<int, string> keyWorddic)
        {

            try
            {
                bool isKeyVerified;
                string url = GetData(KryptonConstants.WHAT);
                if (!keyWorddic.Count.Equals(0))
                {
                    try
                    {
                        IWait<IWebDriver> wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30.00));
                        wait.Until(driver1 => ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete"));
                        SwitchToMostRecentBrowser();
                    }
                    catch
                    {
                        // ignored
                    }
                    isKeyVerified = Utility.DoKeywordMatch(url, Driver.Url);
                }
                else
                {
                    //By Default regular expression match would be done for Url entry in OR.
                    try
                    {
                        IWait<IWebDriver> wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(30.00));
                        wait.Until(driver1 => ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete"));
                        SwitchToMostRecentBrowser();
                    }
                    catch
                    {
                        // ignored
                    }
                    try
                    {
                        Driver.SwitchTo().DefaultContent();
                    }
                    catch (Exception)
                    {
                        //do nothing
                    }
                    Regex regExp = new Regex(url.ToLower());
                    Match m = regExp.Match((Driver.Url).ToLower());
                    isKeyVerified = m.Success;
                }

                if (!isKeyVerified)
                {
                    Property.Remarks = "Page with URL: \"" + url + "\" is not displayed." +
                                              "Actual displayed page was: " + Driver.Url;
                }

                return isKeyVerified;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public Boolean IsAlertPresent()
        {

            Boolean presentFlag = false;

            try
            {

                // Check the presence of alert
                if (Driver != null)
                {
                    IAlert alert = Driver.SwitchTo().Alert();
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
        /// <param name="stepAction"></param>
        /// <returns></returns>
        public string GetScreenShot(string stepActionNumber, string path, string stepAction = "")
        {
            //if alert exists then will not be taken screen shot
            if (IsAlertPresent())
            {
                return string.Empty;
            }
            if (Driver == null)
                return string.Empty;
            Screenshot screenShot = null;
            string imageName = string.Empty;
            string htmlFileName = string.Empty;
            try
            {
                if ((_errorCaptureAs.ToLower().Equals("both") || _errorCaptureAs.ToLower().Equals("image")))
                {

                    SwitchToMostRecentBrowser();

                    //Get current url of the page
                    Property.AttachmentsUrl = Driver.Url;

                    if (IsRemoteExecution.Equals("true"))
                    {
                        try
                        {
                            screenShot = ((ITakesScreenshot)Driver).GetScreenshot();

                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Take snapshot: " + e.Message);
                        }
                    }
                    else
                    {
                        switch (BrowserName.ToLower())
                        {
                            case KryptonConstants.BROWSER_FIREFOX:
                                FirefoxDriver firefoxDriver = (FirefoxDriver)Driver;
                                screenShot = ((ITakesScreenshot)firefoxDriver).GetScreenshot();
                                break;
                            case KryptonConstants.BROWSER_IE:
                                InternetExplorerDriver ieDriver = (InternetExplorerDriver)Driver;
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
                                ChromeDriver chromeDriver = (ChromeDriver)Driver;
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
                        //Saving HQ image first
                        screenShot.SaveAsFile(path + "/" + imageNameHQ, System.Drawing.Imaging.ImageFormat.Jpeg);
                        // Compress Image
                        try
                        {
                            Utility.CompressImage(path + "/" + imageNameHQ, path + "/" + imageName);
                            // Delete HQ image after compression
                            File.Delete(path + "/" + imageNameHQ);
                        }
                        catch
                        {
                            // Ignored
                        }
                    }
                }
                if (_errorCaptureAs.ToLower().Equals("both") || _errorCaptureAs.ToLower().Equals("html"))
                {
                    try
                    {
                        string pageViewSource = Driver.PageSource;
                        htmlFileName = "Html" + stepActionNumber + ".html";
                        StreamWriter sw = new StreamWriter(path + "/" + htmlFileName, false, Encoding.UTF8);
                        //Create html view source file.
                        sw.WriteLine(pageViewSource.ToCharArray());
                        sw.Close();
                    }
                    catch
                    {
                        // Ignored
                    }
                }
                return imageName + "|" + htmlFileName;
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains("No response from server for url."))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }

        /// <summary>
        ///Method to Quit initialized web driver
        /// </summary>
        public void Shutdown()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Driver.Title) && Property.IsSauceLabExecution.Equals(true))
                {
                    //Saucelabs assume driver as diconnected only when using quit().
                    try
                    {
                        Driver.Quit();
                    }
                    catch
                    {
                        // ignored
                    }
                    Driver.Close();
                }
            }
            catch (WebDriverException e)
            {
                if (e.Message.Contains(Exceptions.ERROR_NORESPONSEURL))
                {
                    throw new WebDriverException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0007") + ":" + e.Message);
                }
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0054") + ":" + e.Message);
            }
        }
        /// <summary>
        ///  Method to get the instance of selenium version 1.3
        /// </summary>
        /// <returns>Selenium instance</returns>
        public static Selenium.DefaultSelenium GetSeleniumOne()
        {
            if (_selenium == null)
            {
                _selenium = new Selenium.WebDriverBackedSelenium(Driver, Driver.Url);
            }
            try
            {
                _selenium.Start();
            }
            catch (Exception)
            {
                // ignored
            }
            return _selenium;
        }
    }
}
