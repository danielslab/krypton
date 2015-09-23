/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Krypton.TestEngine.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Implement the Grid functionality.
*****************************************************************************/
using System;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Firefox;
using System.Threading;
using Common;
using System.IO;
using ControllerLibrary;
using System.Xml;

namespace Driver
{
    public class KryptonRemoteDriver : RemoteWebDriver, ITakesScreenshot
    {
        private readonly Uri _remoteHost;

        public KryptonRemoteDriver(Uri remoteHost, ICapabilities capabilities)
            : base(remoteHost, capabilities)
        {
            _remoteHost = remoteHost;

        }

        public static SessionId getExecutionID(IWebDriver driver)
        {
            return ((KryptonRemoteDriver)driver).SessionId;
        }

        public new  Screenshot  GetScreenshot()
        {
            try
            {
                // Get the screenshot as base64.
                Response screenshotResponse = this.Execute(DriverCommand.Screenshot, null);
                string base64 = screenshotResponse.Value.ToString();

                // ... and convert it.
                return new Screenshot(base64);
            }
            catch
            {
                return null;
            }
        }
    }

    public class SeleniumGrid
    {
        private Uri _remoteUrl;
        private DesiredCapabilities _capabilities, s_capabilities;
        private string _browserName;
        private IWebDriver driver;
        private string s_browserName;
        private Uri s_remoteUrl;
        public SeleniumGrid(string remoteUrl, string browser)
        {
            _remoteUrl = new Uri(remoteUrl);
            _browserName = browser;
        }

        //Added for sauce execution 
        public SeleniumGrid(string sremoteUrl, string sbrowser, DesiredCapabilities scapabilities)
        {
            s_remoteUrl = new Uri(sremoteUrl);
            s_browserName = sbrowser;
            s_capabilities = scapabilities;

        }
        //Added for sauce execution 
        public IWebDriver GetDriverSauce()
        {
            s_capabilities.SetCapability(KryptonConstants.BROWSER_NAME, s_browserName);
            driver = new KryptonRemoteDriver(s_remoteUrl, s_capabilities);
            return driver;
        }

       

        public IWebDriver GetDriver(out string hostName)
        {
            hostName = string.Empty;
            switch (_browserName.ToLower())
            {

                case KryptonConstants.BROWSER_FIREFOX:
                    FirefoxProfile remoteProfile = null;
                    try
                    {
                        remoteProfile = RemoteFirefoxProfile();     
                        _capabilities = DesiredCapabilities.Firefox(); 
                        _capabilities.SetCapability(KryptonConstants.BROWSER_NAME, KryptonConstants.BROWSER_FIREFOX);                                
                        _capabilities.SetCapability(KryptonConstants.FIREFOX_PROFILE, remoteProfile.ToBase64String());
                        driver = new KryptonRemoteDriver(_remoteUrl, _capabilities);
                        driver.Manage().Window.Maximize();
                    }
                    catch (Exception ex)
                    {
                        bool IsFFprofileExe = false;
                        if (ex.Message.IndexOf("Access to the path", StringComparison.OrdinalIgnoreCase) >= 0)
                        {

                            IsFFprofileExe = true;
                            try
                            {
                                Thread.Sleep(3000);
                                remoteProfile = null;
                                FirefoxProfile  remoteProfile2 = RemoteFirefoxProfile();
                                _capabilities = null;
                                DesiredCapabilities capabilities = DesiredCapabilities.Firefox();
                                capabilities.SetCapability(KryptonConstants.BROWSER_NAME, KryptonConstants.BROWSER_FIREFOX);
                                capabilities.SetCapability(KryptonConstants.FIREFOX_PROFILE, remoteProfile2.ToBase64String());
                                driver = new KryptonRemoteDriver(_remoteUrl, capabilities);
                                driver.Manage().Window.Maximize();
                            }
                            catch (Exception exmsg)
                            {
                                string errorMsg = string.Empty;
                                while (exmsg != null)
                                {
                                    errorMsg += exmsg.Message + Environment.NewLine;
                                    exmsg = exmsg.InnerException;
                                }
                                throw new Exception("Could Not launch browser: " + errorMsg);

                            }

                        }
                        if (!IsFFprofileExe)
                        {
                            string errorMsg = string.Empty;
                            while (ex != null)
                            {
                                errorMsg += ex.Message + Environment.NewLine;
                                ex = ex.InnerException;
                            }
                            throw new Exception("Could Not launch browser: " + errorMsg);
                        }

                    }
                    if (Browser.IsBrowserDimendion)
                        driver.Manage().Window.Size = new System.Drawing.Size(Browser.width, Browser.height);
                    Browser.driverlist.Add(driver);                   
                    break;
                case KryptonConstants.BROWSER_IE:
                    Thread.Sleep(2000);
                    _capabilities = DesiredCapabilities.InternetExplorer();
                    _capabilities.SetCapability(KryptonConstants.BROWSER_NAME, "internet explorer");
                    //This will ignore protected mode settings check
                    _capabilities.SetCapability("ignoreProtectedModeSettings", true);

                    try
                    {
                        driver = new KryptonRemoteDriver(_remoteUrl, _capabilities);
                        driver.Manage().Window.Maximize();

                    }
                    catch (Exception ex)
                    {
                        string errorMsg = string.Empty;
                        while (ex != null)
                        {
                            errorMsg += ex.Message + Environment.NewLine;
                            ex = ex.InnerException;
                        }
                        throw new Exception("Could Not launch browser: " + errorMsg);

                    }
                    if (Browser.IsBrowserDimendion)
                        driver.Manage().Window.Size = new System.Drawing.Size(Browser.width, Browser.height);
                    Browser.driverlist.Add(driver);
                    break;
                case KryptonConstants.BROWSER_CHROME:                    
                    Thread.Sleep(2000);
                    _capabilities = DesiredCapabilities.Chrome();
                    try
                    {
                        driver = new KryptonRemoteDriver(_remoteUrl, _capabilities); 
                        driver.Manage().Window.Maximize();
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = string.Empty;
                        while (ex != null)
                        {
                            errorMsg += ex.Message + Environment.NewLine;
                            ex = ex.InnerException;
                        }
                        throw new Exception("Could Not launch browser: " + errorMsg);

                    }
                    if (Browser.IsBrowserDimendion)
                        driver.Manage().Window.Size = new System.Drawing.Size(Browser.width, Browser.height);
                    Browser.driverlist.Add(driver);                    
                    break;
                // For Launching Safari on Remote Machine
                case KryptonConstants.BROWSER_SAFARI:
                    Thread.Sleep(2000);
                    _capabilities = DesiredCapabilities.Safari();
                    _capabilities.SetCapability(KryptonConstants.BROWSER_NAME, KryptonConstants.BROWSER_SAFARI);
                    _capabilities.IsJavaScriptEnabled = true;
                    driver = new KryptonRemoteDriver(_remoteUrl, _capabilities);
                    driver.Manage().Window.Maximize();
                    if (Browser.IsBrowserDimendion)
                        driver.Manage().Window.Size = new System.Drawing.Size(Browser.width,Browser.height);
                  Browser.driverlist.Add(driver);
                    break;

            }          
           
           
            try
            {
                SessionId id = KryptonRemoteDriver.getExecutionID(driver);
                hostName = GetNodeHost(_remoteUrl, id);
                try
                {
                   Uri remoteInfo = new Uri(hostName);
                   string ip = remoteInfo.Host.ToString();
                   Property.RemoteMachineIP = ip;
                   hostName=GetMachineNameFromIPAddress(ip);
                   hostName = string.IsNullOrEmpty(hostName) ? ip : hostName;
                    
                }
                catch
                { 
                
                }
            }
            catch{ 
               
                 }
            
            return driver;

        }

        private FirefoxProfile RemoteFirefoxProfile()
        {
            FirefoxProfile remoteProfile = new FirefoxProfile();
            remoteProfile.SetPreference("webdriver_assume_untrusted_issuer", false);
            remoteProfile.AcceptUntrustedCertificates = true;
            remoteProfile.EnableNativeEvents = false;   // For running firefox on linux

            return remoteProfile;
        }

        /// <summary>
        /// </summary>
        /// <param name="_remoteHost"> </param>
        /// <param name="sid"> Current seesion Id</param>
        /// <returns> host name</returns>
        public string GetNodeHost(Uri _remoteHost,SessionId sid)
        {

            var uri = new Uri(string.Format("http://{0}:{1}/grid/api/testsession?session={2}", _remoteHost.Host, _remoteHost.Port, sid));
            var request = (HttpWebRequest)WebRequest.Create(uri);

            request.Method = "POST";

            request.ContentType = "application/json";

            using (var httpResponse = (HttpWebResponse)request.GetResponse())

            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {


                var response = JObject.Parse(streamReader.ReadToEnd());

                return response.SelectToken("proxyId").ToString();

            }

        }


        private string GetMachineNameFromIPAddress(string ipAdress)
        {
            string machineName = string.Empty;
            try
            {
                ServiceAgent service = new ServiceAgent(ipAdress);
                machineName= service.GetMachineName();             
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(machineName);
                machineName= xmlDoc.InnerText;

            }
            catch (Exception ex)
            {
                // Machine not found...
            }
            return machineName;
        }

    }
    
}
