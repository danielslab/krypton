/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Driver.SeleniumPHP.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Class to create Selenium script in PHP language.
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using Common;
using Driver.Browsers;

namespace Driver
{
    public class SeleniumPHP : IScriptSelenium
    {
        public static StringBuilder SeleniumCodeForPhp;
        private Dictionary<string, string> _objDataRow = new Dictionary<string, string>();
        private string _attribute;
        private string _attributeType;
        public SeleniumPHP()
        {
            SeleniumCodeForPhp = new StringBuilder();
            SeleniumCodeForPhp.AppendLine("<?php");
            SeleniumCodeForPhp.AppendLine("require(\"WebDriver/Driver.php\");");
            SeleniumCodeForPhp.AppendLine("require(\"phpunit/PHPUnit/Framework/Assert.php\");");
            InitDriver();
        }

        /// <summary>
        ///  set test object information dictionary.       
        /// </summary>
        /// <param name="objDataRow">Dictionary containing test object information</param>
        public void SetObjDataRow(Dictionary<string, string> objDataRow)
        {
            try
            {
                if (objDataRow.Count > 0)
                {
                    _objDataRow = objDataRow;
                    _attributeType = GetData(KryptonConstants.HOW);
                    _attribute = GetData(KryptonConstants.WHAT);
                }
                else
                {
                    _objDataRow = objDataRow;
                    _attributeType = string.Empty;
                    _attribute = string.Empty;
                }
            }

            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        ///  : This method will provide testobject information and attribute. 
        /// </summary>
        /// <param name="dataType">string : Name of the test object</param>
        /// <returns >string : object information</returns>
        /// 
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
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        ///   : Generate PHP to find element based on property.
        /// </summary>
        public void FindElement()
        {
            _attribute = _attribute.Replace("\"", "\\\"");
            switch (_attributeType.ToLower())
            {
                case "cssselector":
                case "css":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"css=" + _attribute + "\");");
                    break;
                case "name":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"name=" + _attribute + "\");");
                    break;
                case "id":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"id=" + _attribute + "\");");
                    break;
                case "link":
                case "linktext":
                case "text":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"link text=" + _attribute + "\");");
                    break;
                case "xpath":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"xpath=" + _attribute + "\");");
                    break;
                case "partiallinktext":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"partial link text=" + _attribute + "\");");
                    break;
                case "tagname":
                case "tag":
                case "html tag":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"tag name=" + _attribute + "\");");
                    break;
                case "class":
                case "classname":
                    SeleniumCodeForPhp.AppendLine("$webelement=$webdriver->get_element(\"class name=" + _attribute + "\");");

                    break;
            }
        }

        /// <summary>
        ///   : Generate PHP to initialize WebDriver for PHP.
        /// </summary>
        public void InitDriver()
        {
            SeleniumCodeForPhp.AppendLine("$webdriver=WebDriver_Driver::InitAtLocal( \"4444\",\"" + Utility.GetParameter(Property.BrowserString) + "\");");
        }
        /// <summary>
        ///   : Generate PHP to Click on element.
        /// </summary>
        public void Click()
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->click();");
        }
        /// <summary>
        ///   : Generate PHP to type in textbox.
        /// </summary>
        /// <param name="text"></param>
        public void SendKeys(string text)
        {
            text = text.Replace("$", "\\$");
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->send_keys(\"" + text + "\");");
        }
        /// <summary>
        ///   : Generate PHP to check radio-button or checkbox
        /// </summary>
        public void Check()
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->select();");
        }
        /// <summary>
        ///   : Generate PHP to Uncheck radio-button or checkbox.
        /// </summary>
        public void Uncheck()
        {
            //TODO
        }
        /// <summary>
        ///   : Generate PHP to wait WebDriver for a specified time period.
        /// </summary>
        public void Wait()
        {
        }
        /// <summary>
        ///   : Generate PHP to Close WebPage.
        /// </summary>
        public void Close()
        {
            SeleniumCodeForPhp.AppendLine("$webdriver->close_window();");
        }
        /// <summary>
        ///   : Generate PHP to navigate specified URL.
        /// </summary>
        /// <param name="testData">String : URl to navigate</param>
        public void Navigate(string testData)
        {
            if (testData.IndexOf(':') != 1 && !(testData.Contains("http://") || testData.Contains("https://"))) //Check for file protocol and http protocol :
            {
                testData = "http://" + testData;
            }
            SeleniumCodeForPhp.AppendLine("$webdriver->load(\"" + testData + "\");");
        }
        /// <summary>
        ///   : Generate PHP to fire any event.
        /// </summary>
        public void FireEvent()
        {
        }
        /// <summary>
        ///   : Generate PHP to navigate back in web page.
        /// </summary>
        public void GoBack()
        {
            SeleniumCodeForPhp.AppendLine("$webdriver->go_back();");
        }
        /// <summary>
        ///   : Generate PHP to navigate forward in web page.
        /// </summary>
        public void GoForward()
        {
            SeleniumCodeForPhp.AppendLine("$webdriver->go_forward();");
        }
        /// <summary>
        ///   : Generate PHP to refresh web page.
        /// </summary>
        public void Refresh()
        {
            SeleniumCodeForPhp.AppendLine("$webdriver->refresh();");
        }
        /// <summary>
        ///   : Generate PHP to clear text form text box.
        /// </summary>
        public void Clear()
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->clear();");
        }
        /// <summary>
        /// Generate PHP to type unique data in text box.
        /// </summary>
        public void EnterUniqueData()
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->type_random();");
        }
        /// <summary>
        ///   : Generate PHP to enter non alphanumeric key.
        /// </summary>
        /// <param name="testData">String : Non alphanumeric key to enter</param>
        public void KeyPress(string testData = "")
        {
            FindElement();
            switch (testData.ToLower())
            {
                case "enter":
                    SeleniumCodeForPhp.AppendLine("$webelement->send_keys(Keys::ENTER);");
                    break;
                case "space":
                    SeleniumCodeForPhp.AppendLine("$webelement->send_keys(Keys::SPACE);");
                    break;
                case "tab":
                    SeleniumCodeForPhp.AppendLine("$webelement->send_keys(Keys::TAB);");
                    break;
            }
        }
        /// <summary>
        ///   : Generate PHP to select option from the list
        /// </summary>
        /// <param name="testData">string : option to be select</param>
        public void SelectItem(string testData)
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->select_label(\"" + testData + "\");");
            SeleniumCodeForPhp.AppendLine("$option_element=$webelement->get_selected();");
            SeleniumCodeForPhp.AppendLine("$text=$option_element->get_text();");
            SeleniumCodeForPhp.AppendLine("if(strcmp($text," + testData + ")!=0){ \n$webelement->select_value(\"" + testData + "\");\n}");

        }
        /// <summary>
        ///   : Generate PHP to select option from the list based on index
        /// </summary>
        /// <param name="testData">String : Index number of option (1 based)</param>
        public void SelectItemByIndex(string testData)
        {
            testData = (Int32.Parse(testData) - 1).ToString();
            FindElement();
            SeleniumCodeForPhp.AppendLine("$webelement->select_index(\"" + testData + "\");");
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
        }
        public void DismissAlert()
        {
        }
        public void VerifyAlertText()
        {
        }

        /// <summary>
        ///   : Generate PHP to find specified property of Testobject.
        /// </summary>
        /// <param name="property">String : Property Name </param>
        public void GetObjectProperty(string property)
        {

            FindElement();
            string actualPropertyValue = string.Empty;
            string javascript;
            switch (property.ToLower())
            {
                case "text":
                    SeleniumCodeForPhp.AppendLine("$value=$webelement->get_text();");
                    break;
                case "style.backgroundimage":
                    switch (Browser.BrowserName.ToLower())
                    {
                        case KryptonConstants.BROWSER_IE:
                            javascript = "return arguments[0].currentStyle.backgroundImage;";
                            break;
                        case KryptonConstants.BROWSER_FIREFOX:
                            javascript = "return document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"backgroundmage\");";
                            break;
                        default:
                            javascript = "return arguments[0].currentStyle.backgroundImage;";
                            break;
                    }
                    ExecuteScript(javascript, "$webelement");
                    break;

                case "style.color":
                    switch (Browser.BrowserName.ToLower())
                    {
                        case KryptonConstants.BROWSER_IE:
                            javascript = "var color=arguments[0].currentStyle.color;return color;";
                            break;
                        case KryptonConstants.BROWSER_FIREFOX:
                            javascript = "var color= document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"color\");var c = color.replace(/rgb\\((.+)\\)/,'$1').replace(/\\s/g,'').split(',');color = '#'+ parseInt(c[0]).toString(16) +''+ parseInt(c[1]).toString(16) +''+ parseInt(c[2]).toString(16);return color;";
                            break;
                        default:
                            javascript = "var color=arguments[0].currentStyle.color;return color;"; ;
                            break;
                    }
                    ExecuteScript(javascript, "$webelement");
                    break;
                case "style.fontweight":
                case "style.font-weight":
                    switch (Browser.BrowserName.ToLower())
                    {
                        case KryptonConstants.BROWSER_IE:
                            javascript = "var fontWeight= arguments[0].currentStyle.fontWeight;if(fontWeight==700){return 'bold';}if(fontWeight==400){return 'normal';}return fontWait;";
                            break;
                        case KryptonConstants.BROWSER_FIREFOX:
                            javascript = "return document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"font-weight\");";
                            break;
                        default:
                            javascript = "return arguments[0].currentStyle.fontWeight;";
                            break;
                    }
                    ExecuteScript(javascript, "$webelement");

                    break;
                case "tooltip":
                case "title":
                    javascript = "return title=arguments[0].title;";
                    ExecuteScript(javascript, "$webelement");
                    break;
                case "scrollheight":
                    javascript = "return arguments[0].scrollHeight;";
                    ExecuteScript(javascript, "$webelement");
                    break;
                case "file name":
                case "filename":
                    javascript = "var src = arguments[0].src; src=src.split(\"/\"); return(src[src.length-1]);";
                    ExecuteScript(javascript, "$webelement");
                    break;
                case "height":
                    javascript = "return arguments[0].height;";
                    ExecuteScript(javascript, "$webelement");
                    break;
                case "width":
                    javascript = "return arguments[0].width;";
                    ExecuteScript(javascript, "$webelement");
                    break;
                case "checked":
                    javascript = "return arguments[0].checked;";
                    ExecuteScript(javascript, "$webelement");
                    break;
                default:
                    SeleniumCodeForPhp.AppendLine(" $value=$webelement->get_attribute_value(\"" + property + "\");");
                    break;
            }
        }
        public void GetPageProperty(string propertyType)
        {

            switch (propertyType.ToLower())
            {
                case "title":
                    SeleniumCodeForPhp.AppendLine("$propertyval= $webdriver->get_title();");
                    break;
                case "url":
                    SeleniumCodeForPhp.AppendLine("$propertyval= $webdriver->get_url();");
                    break;
            }
        }
        /// <summary>
        ///   : Generate PHP to set property present in DOM
        /// </summary>
        /// <param name="property">string : Name of property</param>
        /// <param name="propertyValue">string : Value of property</param>
        public void SetAttribute(string property, string propertyValue)
        {
            string javascript;
            switch (property.ToLower().Trim())
            {
                case "checked":
                    javascript = "return arguments[0]." + property + " = " + propertyValue + ";";
                    break;
                default:
                    javascript = "return arguments[0]." + property + " = \"" + propertyValue + "\";";
                    break;
            }
            ExecuteScript(javascript);
        }
        /// <summary>
        ///   : Generate PHP to Verify Page Property 
        /// </summary>
        /// <param name="property">String : Property Name</param>
        /// <param name="propertyValue">String : Property Value to be verify</param>
        public void VerifyPageProperty(string property, string propertyValue)
        {
            GetPageProperty(property);
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertEquals(\"" + propertyValue + "\",propertyval,\"Values mismatched\");");

        }
        /// <summary>
        ///   : Generate PHP to Verify that webpage is displayed correctely or not.
        /// </summary>
        public void VerifyPageDisplay()
        {
            string url = GetData(KryptonConstants.WHAT);
            GetPageProperty("url");
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertEquals(\"" + url + "\",propertyval,\"Page not display\");");
        }

        /// <summary>
        ///   : Generate PHP to verify text present on page.
        /// </summary>
        /// <param name="text">String : Text to be verify</param>
        public void VerifyTextPresentOnPage(string text)
        {
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertContains(\"" + text + "\",$webdriver->get_text(),\"Text mismatched\");");
        }
        /// <summary>
        ///   : Generate PHP to verify text not present on page.
        /// </summary>
        /// <param name="text">String : text to be verify</param>
        public void VerifyTextNotPresentOnPage(string text)
        {
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertNotContains(\"" + text + "\",$webdriver->get_text(),\"Text mismatched\");");
        }
        /// <summary>
        ///   : Generate PHP to verify item present in list.
        /// </summary>
        /// <param name="listItem">String : list item to be verify</param>
        public void VerifyListItemPresent(string listItem)
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine(" $options=$webelement->get_options();\n$status=false;\nforeach($options as &$option)\n{\n	$text=$option->get_text();\n");
            SeleniumCodeForPhp.AppendLine("if(strcmp($text,\"" + listItem + "\")==0)\n	{\n		$status=true;\n		break;\n	}\n	if(!$status)\n	{");
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertTrue(false,\"Validation fail\");\n}\n  }");
        }
        /// <summary>
        ///   : Generate PHP to verify item not present in list.
        /// </summary>
        /// <param name="listItem">String : list item to be verify</param>
        public void VerifyListItemNotPresent(string listItem)
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine(" $options=$webelement->get_options();\n$status=false;\nforeach($options as &$option)\n{\n	$text=$option->get_text();\n");
            SeleniumCodeForPhp.AppendLine("if(strcmp($text,\"" + listItem + "\")==0)\n	{\n		PHPUnit_Framework_Assert::assertTrue(false,\"Validation fail\");		}\n}");

        }
        /// <summary>
        ///   : Generate PHP to verify test object display correctly or not.
        /// </summary>
        public void VerifyObjectPresent()
        {
            VerifyObjectDisplay();
        }
        /// <summary>
        ///   : Generate PHP to verify test object not display.
        /// </summary>
        public void VerifyObjectNotPresent()
        {
            VerifyObjectNotDisplay();
        }
        /// <summary>
        /// Generate PHP to verify object property.
        /// </summary>
        /// <param name="property">String : property name</param>
        /// <param name="propertyVal">String : property value</param>
        public void VerifyObjectProperty(string property, string propertyVal)
        {
            GetObjectProperty(property);
            SeleniumCodeForPhp.AppendLine("if(strcmp($value,\"" + propertyVal + "\")!=0){PHPUnit_Framework_Assert::assertTrue(false,\"Validation fail\");\n}");

        }

        /// <summary>
        ///   :  Generate PHP to verify test object property does not match.
        /// </summary>
        /// <param name="property">String : property name </param>
        /// <param name="propertyVal">String : property value</param>
        public void VerifyObjectPropertyNot(string property, string propertyVal)
        {
            GetObjectProperty(property);
            SeleniumCodeForPhp.AppendLine("if(strcmp($value,\"" + propertyVal + "\")==0){PHPUnit_Framework_Assert::assertTrue(false,\"Validation fail\");\n}");
        }

        /// <summary>
        ///   : Generate PHP to verify text present on web page view source.
        /// </summary>
        /// <param name="text">String: text to be verify</param>
        public void VerifyTextInPageSource(string text)
        {
            SeleniumCodeForPhp.AppendLine("$strsource=$webdriver->get_source();");
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertContains(\"" + text + "\",$strsource,\"text not contained\",true);");
        }
        /// <summary>
        ///  Boara : Generate PHP to verify text not present on web page view source.
        /// </summary>
        /// <param name="text">String : text to be verify</param>
        public void VerifyTextNotInPageSource(string text)
        {
            SeleniumCodeForPhp.AppendLine("$strsource=$webdriver->get_source();");
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertNotContainsContains(\"" + text + "\",$strsource,\"text not contained\",true);");
        }

        /// <summary>
        ///   : Generate PHP to verify test object display on web page.
        /// </summary>
        public void VerifyObjectDisplay()
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertTrue($webelement->is_visible(), \"Failed asserting that <{$this->locator}> is visible.\");");
        }
        /// <summary>
        ///  :  Generate PHP to verify test object doesn't display on web page.;
        /// </summary>
        public void VerifyObjectNotDisplay()
        {
            FindElement();
            SeleniumCodeForPhp.AppendLine("PHPUnit_Framework_Assert::assertFalse($webelement->is_visible(), \"Failed asserting that <{$this->locator}> is not visible.\");");
        }
        /// <summary>
        ///   : Generate PHP to execute java script.
        /// </summary>
        /// <param name="javascript">String : javascript to be execute</param>
        /// <param name="testobject">String : test object name</param>
        public void ExecuteScript(string javascript, string testobject = "")
        {
            javascript = javascript.Replace("\"", "\\\"");
            SeleniumCodeForPhp.AppendLine("$javascript=\"" + javascript + "\";");
            if (testobject == string.Empty)
            {
                testobject = "\"\"";
            }
            SeleniumCodeForPhp.AppendLine("$value=$webdriver->execute_js_sync($javascript," + testobject + ");");
        }
        /// <summary>
        ///   : Generate PHP to delete cookies of current web page.
        /// </summary>
        /// <param name="testData">String : URL of webpage</param>
        public void DeleteCookies(string testData)
        {
            SeleniumCodeForPhp.AppendLine("$webelement->delete_all_cookies();");
        }
        /// <summary>
        ///   : Generate PHP to quit web driver.
        /// </summary>
        public void ShutDownDriver()
        {

            SeleniumCodeForPhp.AppendLine("$webdriver->quit();");
        }
        /// <summary>
        ///   :Methods to save PHP script on file system.
        /// </summary>
        public void SaveScript()
        {
           
        }
    }
}
