using System;
using System.Collections.Generic;
using OpenQA.Selenium.IE;
using Common;
using OpenQA.Selenium;

namespace Driver.Browsers
{
    class IE
    {
        internal void IntializeDriver(ref IWebDriver driver, ref bool IsBrowserDimendion, ref List<IWebDriver> driverlist, ref int width, ref int height,ref bool DeleteCookie)
        {
            InternetExplorerOptions options = new InternetExplorerOptions();
            // Commented to check CSA related problem
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            options.EnablePersistentHover = false; //added for IE-8 certificate related issue.
            if (DeleteCookie)
            {
                options.EnsureCleanSession = true;
            }
            driver = new InternetExplorerDriver(Property.ApplicationPath + @"\Exes", options);
            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
            driver.Manage().Window.Maximize();    //  Added to maximize IE window forcibely, as this code is updated Action file.                                        
            if (IsBrowserDimendion)
                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
            driverlist.Add(driver);
        }


        /// <summary>
        /// Handle Certificate Navigation block in IE.
        /// </summary>
        public static void handleIECertificateError(ref IWebDriver driver)
        {
            try
            {
                if (driver.Title.Contains("Certificate Error: Navigation Blocked"))
                {
                    driver.Navigate().GoToUrl("javascript:document.getElementById('overridelink').click()");
                }
            }
            catch (Exception e)
            {
                KryptonException.Writeexception(e);
            }

        }
    }
}
