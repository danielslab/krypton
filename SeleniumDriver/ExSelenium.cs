/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: ExSelenium.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Contains Methods For the Element Operations.
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace Driver
{
    class ExSelenium
    {
        public static void WaitForElement(By SearchBy, int TimeoutSeconds, bool IsIgnoreElementVisibilty)
        {
            try
            {
                Browser.switchToMostRecentBrowser();
                WebDriverWait wait = new WebDriverWait(TestObject.driver, TimeSpan.FromSeconds(TimeoutSeconds));
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                if (IsIgnoreElementVisibilty)
                    wait.Until(ExpectedConditions.ElementExists(SearchBy));
                else
                    wait.Until(ExpectedConditions.ElementIsVisible(SearchBy));

            }            
            catch (Exception)
            { }

        }

        public static By GetSelectionMethod(string FindBy, string Locator)
        {
            try
            {
                switch (FindBy.ToLower())
                {
                    case "cssselector":
                        return (By.CssSelector(Locator));
                    case "name":
                        return (By.Name(Locator));
                    case "id":
                        return (By.Id(Locator));
                    case "linktext":
                    case "text":
                    case "link":
                        return (By.LinkText(Locator));
                    case "xpath":
                        return (By.XPath(Locator));
                    case "partiallinktext":
                        return (By.PartialLinkText(Locator));
                    case "tagname":
                    case "htmltag":
                        return (By.TagName(Locator));
                    case "classname":
                        return (By.ClassName(Locator));
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
