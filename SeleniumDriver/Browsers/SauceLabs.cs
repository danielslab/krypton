using System;
using System.Text;
using Common;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using System.Net;
using System.IO;
using System.Diagnostics;

namespace Driver.Browsers
{
    class SauceLabs
    {
        private string username = string.Empty;
        private string accessKey = string.Empty;

        internal void IntializeDriver(ref string remoteUrl,ref string browserName, ref IWebDriver driver, ref Actions driverActions)
        {
            if (Property.RemoteUrl.ToLower().Contains("saucelabs"))
            {
                DesiredCapabilities capabilities = new DesiredCapabilities();
                Utility.SetParameter("CloseBrowserOnCompletion", "true");//Forcing the close browser to true.
                Utility.SetVariable("CloseBrowserOnCompletion", "true");
                Property.IsSauceLabExecution = true;
                capabilities.SetCapability("username", Utility.GetParameter("username"));//Registered user name of Sauce labs
                capabilities.SetCapability("accessKey", Utility.GetParameter("password"));// Accesskey provided by the Sauce labs 
                capabilities.SetCapability("platform", Utility.GetParameter("Platform"));// OS on which execution is to be done Eg: Windows 7 , mac , Linux etc..
                capabilities.SetCapability("name", Utility.GetParameter("TestCaseId"));
                capabilities.SetCapability("browser", Utility.GetParameter("SauceBrowser"));
                capabilities.SetCapability("version", Utility.GetParameter("VersionofBrowser"));
                string RemoteHost = string.Empty;
                remoteUrl = Property.RemoteUrl + "/wd/hub";

                // if Sauce connect is required...
                string isSauceConnectRequired = Utility.GetParameter("IsTestEnvironment");
                if (isSauceConnectRequired.ToLower() == "true")
                {
                    ExecuteSauceConnect();
                    Thread.Sleep(20 * 1000);
                }

                SeleniumGrid oSeleniumGrid = new SeleniumGrid(remoteUrl, Utility.GetParameter("SauceBrowser"), capabilities);
                Browser.Driver = oSeleniumGrid.GetDriverSauce();
                if (!string.IsNullOrEmpty(RemoteHost))
                {
                    Property.RcMachineId = RemoteHost;
                    Utility.SetVariable(Property.RcMachineId, RemoteHost);
                    Utility.SetParameter(Property.RcMachineId, RemoteHost);
                }
                driverActions = new Actions(driver);

            }
            else
            {
                string RemoteHost = string.Empty;
                remoteUrl = Property.RemoteUrl + "/wd/hub";
                // (management of remote driver in a seperate class)
                SeleniumGrid oSeleniumGrid = new SeleniumGrid(remoteUrl, browserName);
                Browser.Driver = oSeleniumGrid.GetDriver(out RemoteHost);
                if (!string.IsNullOrEmpty(RemoteHost))
                {
                    Property.RcMachineId = RemoteHost;
                    Utility.SetVariable(Property.RcMachineId, RemoteHost);
                    Utility.SetParameter(Property.RcMachineId, RemoteHost);
                }
                //Initializing actions object for later usage
                driverActions = new Actions(driver);
            }
        }

        internal void UpdateJobDeatils()
        {
            StreamReader reader = null;
            try
            {
                HttpWebRequest request = WebRequest.Create("https://" + username + ":" + accessKey + "@saucelabs.com/rest/v1/" + username + "/jobs/" + Utility.GetVariable("SessionID")) as HttpWebRequest;
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
                reader = new StreamReader(response.GetResponseStream());
                string contents = reader.ReadLine();

            }
            catch (Exception ex)
            {
                KryptonException.Writeexception(ex);
            }
            finally 
            {
                if (reader != null) reader.Close();
            }
        }

        public static void ExecuteSauceConnect()
        {
            try
            {

                foreach (Process proc in Process.GetProcessesByName("sc"))
                {
                    proc.Kill();
                }
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                startInfo.FileName = Utility.GetParameter("SauceConnectPath");
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = "-u " + Utility.GetParameter("username") + " " + "-k " + Utility.GetParameter("password") + Utility.GetParameter("CommandLineOptions");
                using (Process exeProcess = Process.Start(startInfo)) ;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }
    }
}
