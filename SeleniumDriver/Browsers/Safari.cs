using System;
using System.Collections.Generic;
using OpenQA.Selenium.Safari;

namespace Driver.Browsers
{
    class Safari
    {
        internal void IntializeDriver(ref OpenQA.Selenium.IWebDriver driver, ref bool IsBrowserDimendion, ref List<OpenQA.Selenium.IWebDriver> driverlist, ref int width, ref int height)
        {
            driver = new SafariDriver();
            driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(150));
            driver.Manage().Window.Maximize();
            if (IsBrowserDimendion)
                driver.Manage().Window.Size = new System.Drawing.Size(width, height);
            driverlist.Add(driver);
        }
    }
}
