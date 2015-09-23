/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Driver.SeleniumJava.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Test Object action class
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Common;

namespace Driver
{
    public class SeleniumJava : IScriptSelenium
    {
        public static StringBuilder seleniumCodeForJava;
        private Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        public SeleniumJava()
        {
            seleniumCodeForJava = new StringBuilder();
            seleniumCodeForJava.AppendLine("import java.util.List;");
            seleniumCodeForJava.AppendLine("import java.util.*;");
            seleniumCodeForJava.AppendLine("import java.text.SimpleDateFormat;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.By;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.RenderedWebElement;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.WebDriver;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.WebElement;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.firefox.*;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.ie.*;");
            seleniumCodeForJava.AppendLine("import org.openqa.selenium.chrome.*;");
            seleniumCodeForJava.AppendLine("public class Selenium {");
            seleniumCodeForJava.AppendLine("private static WebDriver driver;");
            seleniumCodeForJava.AppendLine("private static WebElement testObject;");
            seleniumCodeForJava.AppendLine("public static void main(String[] args) { \ntry\n{");
            this.InitDriver();
        }

        public void InitDriver()
        {
            switch (Common.Utility.GetParameter(Common.Property.BrowserString))
            {
                case KryptonConstants.BROWSER_FIREFOX:
                    seleniumCodeForJava.AppendLine("driver = new FirefoxDriver();");
                    break;
                case KryptonConstants.BROWSER_IE:
                    seleniumCodeForJava.AppendLine("driver = new InternetExplorerDriver();");
                    break;
                case KryptonConstants.BROWSER_CHROME:
                    seleniumCodeForJava.AppendLine("driver = new ChromeDriver();");
                    break;
            }
        }
        public void SetObjDataRow(Dictionary<string, string> objDataRow)
        {
        }
        public void FindElement()
        {
            try
            {
                string attribute = objDataRow[KryptonConstants.WHAT];
                attribute = attribute.Replace("\"", "\\\"");
                switch (objDataRow[KryptonConstants.HOW].ToLower())
                {
                    case "cssselector":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.cssSelector(\"" + attribute + "\"));");
                        break;
                    case "css":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.cssSelector(\"" + attribute + "\"));");
                        break;
                    case "name":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.name(\"" + attribute + "\"));");
                        break;
                    case "id":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.id(\"" + attribute + "\"));");
                        break;
                    case "linktext":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.linkText(\"" + attribute + "\"));");
                        break;
                    case "link":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.linkText(\"" + attribute + "\"));");
                        break;
                    case "xpath":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.xpath(\"" + attribute + "\"));");
                        break;
                    case "partiallinktext":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.partialLinkText(\"" + attribute + "\"));");
                        break;
                    case "tagname":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.tagName(\"" + attribute + "\"));");
                        break;
                    case "tag":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.tagName(\"" + attribute + "\"));");
                        break;
                    case "html tag":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.tagName(\"" + attribute + "\"));");
                        break;
                    case "classname":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.className(\"" + attribute + "\"));");
                        break;
                    case "class":
                        seleniumCodeForJava.AppendLine("testObject = driver.findElement(By.className(\"" + attribute + "\"));");
                        break;
                }
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {

            }
            catch (Exception e)
            {

            }

        }
        public void KeyPress()
        {
        }

        public void Click()
        {
            seleniumCodeForJava.AppendLine("testObject.click();"); //for java code
        }
        public void SendKeys(string text)
        {
        }
        public void Check()
        {
            seleniumCodeForJava.AppendLine("testObject.setSelected();"); //for java code
        }
        public void Uncheck()
        {
        }
        public void Wait()
        {
        }
        public void Close()
        {
        }
        public void Navigate(string TestData)
        {
            seleniumCodeForJava.AppendLine("driver.get(\"" + TestData + "\");");
        }
        public void DeleteCookies(String TestData)
        {
            seleniumCodeForJava.AppendLine("driver.manage().deleteAllCookies();");//for java
            seleniumCodeForJava.AppendLine("driver.get(\"" + TestData + "\");");//for java . 
        }

        public void FireEvent()
        {
        }
        public void GoBack()
        {
            seleniumCodeForJava.AppendLine("driver.navigate().back();");//for java code
        }
        public void GoForward()
        {
            seleniumCodeForJava.AppendLine("driver.navigate().forward();");//for java code
        }
        public void Refresh()
        {
            seleniumCodeForJava.AppendLine("driver.navigate().refresh();");//for java code
        }
        public void Clear()
        {
            seleniumCodeForJava.AppendLine("testObject.clear();"); //for java code
        }
        public void EnterUniqueData()
        {
            seleniumCodeForJava.AppendLine("testObject.sendKeys(new SimpleDateFormat(\"yyyyMMMddHHmmss\").format(Calendar.getInstance()));"); //for java code
        }
        public void KeyPress(string TestData)
        {
            switch (TestData.ToLower())
            {
                case "arrowdown":
                    seleniumCodeForJava.AppendLine("testObject.sendKeys(Keys.DOWN);"); //for java code
                    break;
                case "enter":
                    seleniumCodeForJava.AppendLine("testObject.sendKeys(Keys.ENTER);"); //for java code
                    break;
            }
        }
        public void SelectItem(string TestData)
        {
            seleniumCodeForJava.AppendLine("new Select(testObject).selectByVisibleText(" + TestData + ");");//for java code
        }
        public void SelectItemByIndex(string TestData)
        {
            seleniumCodeForJava.AppendLine("new Select(testObject).selectByIndex(" + Int32.Parse(TestData) + ");");//for java code
        }
        public void WaitForObject()
        {
        }
        public void WaitForObjectNotPresent()
        {
        }
        public void WaitForObjectProperty()
        {
        }
        public void AcceptAlert()
        {
            seleniumCodeForJava.AppendLine("driver.switchTo().alert().accept();");// for java code
        }
        public void DismissAlert()
        {
            seleniumCodeForJava.AppendLine("driver.switchTo().alert().dismiss();");// for java code
        }
        public void VerifyAlertText()
        {
        }
        public void GetObjectProperty(string property)
        {
        }
        public void GetPageProperty(string propertyType)
        {
        }
        public void SetAttribute(string property, string propertyValue)
        {
        }
        public void VerifyPageProperty(string property, string propertyVal)
        {
        }
        public void VerifyTextPresentOnPage(string text)
        {
        }
        public void VerifyTextNotPresentOnPage(string text)
        {
        }
        public void VerifyListItemPresent(string listItem)
        {
        }
        public void VerifyListItemNotPresent(string listItem)
        {
        }
        public void VerifyObjectPresent()
        {
        }
        public void VerifyObjectNotPresent()
        {
        }
        public void VerifyObjectProperty(string property, string propertyVal)
        {
        }
        public void VerifyObjectPropertyNot(string property, string propertyVal)
        {
        }
        public void VerifyPageDisplay()
        {
        }
        public void VerifyTextInPageSource(string text)
        {
        }
        public void VerifyTextNotInPageSource(string text)
        {
        }
        public void VerifyObjectDisplay()
        {
        }
        public void VerifyObjectNotDisplay()
        {
        }

        public void ExecuteScript(string javascript, string testobject = "")
        {
        }
        public void ShutDownDriver()
        {
            seleniumCodeForJava.AppendLine("driver.quit();");//for java code

        }
        public void SaveScript()
        {
        }

    }
}
