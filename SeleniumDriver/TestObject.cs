/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Driver.TestManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Test Object action class
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Safari;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using Common;
using System.IO;


namespace Driver
{
    public class TestObject : ITestObject
    {
        public static IWebDriver driver;
        private IWebElement testObject;
        private IWebElement firstObject;    //This will contain first object from collection
        private System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> testObjects;
        public static string attributeType;
        public static string attribute;
        private string property;
        private string propertyValue;
        private Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        public IWebElement frameObject;
        //Dictionary to store values from mapping column of object repository
        private Dictionary<string, string> dicMapping = new Dictionary<string, string>();
        private Dictionary<int, string> optiondic = null;
        private Selenium.DefaultSelenium selenium = null;
        public string DebugMode = Common.Property.DebugMode;
        public static WebDriverWait objloadingWait;
        private List<string> stepsForIgnoreWait = new List<string>() {"settestmode","verifyobjectnotpresent","verifyobjectnotdisplayed" };
        public Func<IWebDriver, bool> delIsObjectLoaded;
        private string modifier = string.Empty;

        public TestObject(string waitTimeForObject = null)
        {
            driver = Driver.Browser.driver;
            if (waitTimeForObject != null && driver != null)
            {
                objloadingWait = new WebDriverWait(driver, TimeSpan.FromSeconds(Double.Parse(waitTimeForObject)));
            }
        }


        /// <summary>
        /// set test object information dictionary.
        /// </summary>
        /// <param name="objDataRow">Dictionary containing test object information</param>
        public void SetObjDataRow(Dictionary<string, string> objDataRow, string CurrentStepAction="")
        {
            
                dicMapping.Clear();
                if(objDataRow.Count > 0)
                {
                    this.objDataRow = objDataRow;
                    attributeType = this.GetData(KryptonConstants.HOW);
                    attribute = this.GetData(KryptonConstants.WHAT);

                    //Retrive mapping column from object repository and parse
                    string mapping = objDataRow[KryptonConstants.MAPPING];
                    if (!mapping.Equals(string.Empty))
                    {
                        Array arrMapping = mapping.Split('|');
                        string mapName = string.Empty;
                        string mapValue = string.Empty;
                        foreach (string mappingPair in arrMapping)
                        {
                            mapName = mappingPair.Split('=')[0].Trim();
                            mapValue = mappingPair.Split('=')[1].Trim();
                            dicMapping.Add(mapName, mapValue);
                        }
                    }

                }
                else
                {
                    this.objDataRow = objDataRow;
                    attributeType = string.Empty;
                    attribute = string.Empty;
                }
                RecoveryScenarios.cacheAttribute(attributeType, attribute);
                //wait for object
                try
                {
                    By obBy = ExSelenium.GetSelectionMethod(attributeType, attribute);
                    if (obBy != null)
                    {
                        if (stepsForIgnoreWait.Contains(CurrentStepAction.ToLower()))
                            ExSelenium.WaitForElement(obBy, 02, true);
                        else
                            ExSelenium.WaitForElement(obBy, 10, true);
                    }

                }
                catch(Exception e) 
                {
                    throw e;
                }
            
        }

        //When Object Definition is embedded in the test case sheet instead of OR.
        public void SetObjDataRow(String testObject)
        {
            try
            {
                if (testObject != null && testObject.Split(':').Length>=2)
                {
                    attributeType = testObject.Split(':')[0].Trim();
                    attribute = testObject.Split(':')[1].Trim();
                }
                else
                {
                    attributeType = string.Empty;
                    attribute = string.Empty;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /// <summary>
        /// This method will provide testobject information and attribute. 
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
                    case KryptonConstants.PARENT:
                        return objDataRow[KryptonConstants.PARENT];
                    case KryptonConstants.TEST_OBJECT:
                    case "testobject":
                    case "child":
                        return objDataRow[KryptonConstants.TEST_OBJECT];
                    case KryptonConstants.LOGICAL_NAME:
                        return objDataRow[KryptonConstants.LOGICAL_NAME];
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
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                throw new System.Collections.Generic.KeyNotFoundException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0008"));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Wait For an Element for specified time in seconds and return that element.
        /// </summary>
        /// <returns>IWebElement : element to perform action</returns>
        private IWebElement WaitAndGetElement()
        {
            try
            {
                //Disabling wait for any object in case of browser recovery.
                if (Common.Property.isRecoveryRunning || Common.Property.NoWait == true)
                {
                    this.GetElement(driver);
                }
                else
                {
                    Func<IWebDriver, IWebElement> delIsObjectLoaded;
                    delIsObjectLoaded = this.GetElement;
                    objloadingWait.Until(delIsObjectLoaded);
                }

                //Return first object if no object could be located, and no visible objects were still present
                if ((testObjects.Count() >= 1)
                        && !Common.Property.isRecoveryRunning
                        && testObject == null)
                    {
                        testObject = testObjects.First();
                    }
                return testObject;

            }
            catch (WebDriverException e)
            {
                if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // updated By 
                else
                    if (attribute.Equals(String.Empty) && attributeType.Equals(String.Empty))
                        throw new Exception("Could not found object :  " + e.Message);
                else
                    throw new NoSuchElementException(Common.exceptions.ERROR_OBJECTNOTFOUND+"{ method: "+ attributeType + ",  selector: "+ attribute+ "  }"); 
                   
            }
            catch (System.Collections.Generic.KeyNotFoundException kn)
            {
                if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // updated By 
                else
                    if (attribute.Equals(String.Empty) && attributeType.Equals(String.Empty))
                        throw new Exception("Could not found object :  " + kn.Message);
                    else
                        throw new NoSuchElementException(Common.exceptions.ERROR_OBJECTNOTFOUND+"{ method: " + attributeType + ",  selector: " + attribute + "  }"); 
            }
            catch (System.TimeoutException te)
            {
                if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // updated By 
                else
                    if (attribute.Equals(String.Empty) && attributeType.Equals(String.Empty))
                        throw new Exception("Could not found object :  " + te.Message);
                    else
                        throw new NoSuchElementException(Common.exceptions.ERROR_OBJECTNOTFOUND + " { method: " + attributeType + ",  selector: " + attribute + "  }"); 
                   
            }
            catch (Exception e)
            {
                if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // updated By 
                else
                    if (attribute.Equals(String.Empty) && attributeType.Equals(String.Empty))
                        throw new Exception("Could not found object :  " + e.Message);
                    else
                        throw new NoSuchElementException(Common.exceptions.ERROR_OBJECTNOTFOUND +" { method: " + attributeType + ",  selector: " + attribute + "  }"); 
                              
            }
        }
          
        /// <summary>
        /// Function verify the sort order of the contents in specified test element.
        /// </summary>
        /// <param name="propertyForSorting">string : Property for extacting element's content.</param>
        /// <param name="sortOrder">string : Sort Order</param>
        public bool verifySortOrder(string propertyForSorting, string sortOrder = "")
        {
            char splitCriteria = '\n';
            string textContent = string.Empty;
            string[] lstOfContents = null;
            bool result = false;
            try
            {
                testObject = this.WaitAndGetElement();
                switch (propertyForSorting.ToLower())
                {
                    case "text":
                        textContent = testObject.Text;
                        lstOfContents = textContent.Split(splitCriteria);
                        break;
                    default:
                        break;
                }

                result = this.isSorted(lstOfContents, sortOrder);
            }
            catch (Exception e)
            {
                throw e;
            }
            return result;
        }

        /// <summary>
        /// Determine the string array is sorted in specified order or not.
        /// </summary>
        /// <param name="strArray">string[] : string array </param>
        /// <param name="sortOrder">string : sorting order.</param>
        /// <returns></returns>
        private bool isSorted(string[] strArray, string sortOrder = "")
        {
            try
            {
                if (sortOrder.ToLower().Equals("desc"))
                {
                    for (int i = strArray.Length - 2; i >= 0; i--)
                    {
                        if (strArray[i].CompareTo(strArray[i + 1]) < 0)
                        {
                            return false;
                        }
                    }
                    return true;
                }
                else // By default check for ascending order.
                {
                    for (int i = 1; i < strArray.Length; i++)
                    {
                        if (strArray[i - 1].CompareTo(strArray[i]) > 0)
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }
            catch (System.NullReferenceException)
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0056"));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Wait for Element not displayed on page.
        /// </summary>
        /// <returns>return Element.</returns>
        private bool WaitNonPresenceAndGetElement()
        {
            try
            {
                delIsObjectLoaded = this.VerifyObjectNotDisplayedCondition;
                objloadingWait.Until(delIsObjectLoaded);
                return true;
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf(Common.exceptions.ERROR_OBJECTREPOSITORY, StringComparison.OrdinalIgnoreCase) >= 0)
                    throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0051").Replace("{MSG}", e.Message));
                else
                    throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0020"));
            }
        }


        /// <summary>
        ///Search frames in webpage if found search testobject in searched frames.If no frame found search testobject direct into 
        /// </summary>
        /// <param name="attributeType">string : Property name</param>
        /// <param name="attribute">string : Property value</param>
        /// <returns>IwebElement Instance</returns>
        private IWebElement GetElement(IWebDriver driver)
        {
            try
            {
                #region search element direct to web page.
                if (attribute.Equals(String.Empty) && attributeType.Equals(String.Empty))
                    throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0021"));

                //Make a default attempt, but it might throw errors in case previous action have closed active browser
                IWebElement element = null;
                element = this.GetElementByAttribute(attributeType, attribute);
                if (element != null)
                {
                    return element;
                }
                #endregion

                #region switch to most recent browser and handle IE certificate warning and try again
                try
                {
                    Driver.Browser.switchToMostRecentBrowser();
                    element = this.GetElementByAttribute(attributeType, attribute);
                    if (element != null)
                    {
                        return element;
                    }
                }
                catch(Exception e)
                {
                    Common.KryptonException.writeexception(e);
                }
                #endregion

                #region recover browser and try again
                 //Temporarily commented as it is creating wrong attributes
                if (!Common.Property.isRecoveryRunning)
                    RecoveryScenarios.recoverFromBrowsers();
                element = this.GetElementByAttribute(attributeType, attribute);
                if (element != null)
                {
                    return element;
                }         
                #endregion

                #region search element in frames
                string frame = this.GetData(KryptonConstants.MAPPING);

                if (!frame.Equals(string.Empty))
                {
                    string[] keywords = frame.Split('=');
                    if (keywords[0].ToLower().Trim().Equals("frame"))
                    {
                        frame = keywords[1].Trim();
                    }
                    else
                    {
                        frame = string.Empty;
                    }
                }

                switch (frame.Equals(string.Empty))
                {
                    case true:
                        // Handles sub frames.
                        frameObject = null;
                        getElementFromFrames();
                        if (frameObject != null)
                        {
                            return frameObject;
                        }                        
                        throw new NoSuchElementException();                   
                    case false:
                        driver.SwitchTo().Frame(frame);
                        frameObject = this.GetElementByAttribute(attributeType, attribute);
                        if (frameObject != null)
                        {
                            return frameObject;
                        }
                        throw new NoSuchElementException();

                    default:
                        return null;
                #endregion
                }
            }
            catch (NoSuchElementException)
            {
               // Call recovery.
                if (!Common.Property.isRecoveryRunning && !Common.Property.NoWait)
                {
                    RecoveryScenarios.recoverFromBrowsers();
                    IWebElement element = this.GetElementByAttribute(attributeType, attribute);

                    if (element != null)
                        return element;                  
                }
                return null;
            }
            catch (NoSuchFrameException)
            {
                throw;
            }
            //StaleElementReferenceException is the WebDriver Exception, it usually throw when we try to switch 
            // to a frame and it is not loaded properly(as webdriver only wait for page load not for frame load).
            catch (OpenQA.Selenium.StaleElementReferenceException e)
            {
                return null;
            }
            catch (WebDriverException e)
            {
                throw e;
            }
            catch (Exception e)
            {
               if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // updated By 
                 else
                      throw;
            }
        }

        /// <summary>
        ///Get Frames collection in which the driver is in focus.
        /// </summary>
        /// <returns>IEnumerable<IWebElement> collection</returns>
        private IEnumerable<IWebElement> getFrameElement()
        {

            IEnumerable<IWebElement> frames = driver.FindElements(By.TagName("frame"));
            IEnumerable<IWebElement> iframes = driver.FindElements(By.TagName("iframe"));
            IEnumerable<IWebElement> openiframe = driver.FindElements(By.CssSelector("#open_iframe"));
            IEnumerable<IWebElement> elements = openiframe.Union(frames.Union(iframes));
            return elements;
        }

        /// <summary>
        ///Iterate through all sub frames to get testObject.
        /// </summary>
        /// <returns>IWebElement : TestObject found.</returns>
        private void getElementFromFrames()
        {
            IEnumerable<IWebElement> elements;
            try
            {
                elements = getFrameElement();
            }
            catch (System.InvalidOperationException)
            {
                throw new NoSuchElementException();
            }
            IWebElement testobject;
            if (elements != null && elements.Count() > 0)
            {
                foreach (IWebElement element1 in elements)
                {
                    if (frameObject != null)
                    {
                        break;
                    }
                    try
                    {
                        driver.SwitchTo().Frame(element1);
                        
                        testobject = this.GetElementByAttribute(attributeType, attribute);
                        
                        if (testobject != null)
                        {
                            frameObject = testobject;
                            break;
                        }
                    }
                    catch
                    {
                        // Sometime, when switching to a frame, staleElement exception occures
                        //so switching to default contents becomes necessary before switching to frame/ iFrame in question
                        try
                        {
                            driver.SwitchTo().DefaultContent();
                            Thread.Sleep(100);
                            driver.SwitchTo().Frame(element1);
                            testobject = this.GetElementByAttribute(attributeType, attribute);

                            if (testobject != null)
                            {
                                frameObject = testobject;
                                break;
                            }
                            else
                            {
                                getElementFromFrames();
                            }
                        }
                        catch (Exception ex)
                        {
                            string sd = ex.Message;
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Create IWebElement using  Element attributes type and Element attribute.
        /// </summary>
        /// <param name="attributeType">string : Type of property</param>
        /// <param name="attribute">string : Property value</param>
        /// <returns>Instance of IwebElement</returns>
        private IWebElement GetElementByAttribute(string attributeType, string attribute)
        {
            try
            {
                switch (attributeType.ToLower())
                {
                    case "cssselector":
                        testObjects = driver.FindElements(By.CssSelector(attribute));
                        break;
                    case "css":
                        testObjects = driver.FindElements(By.CssSelector(attribute));
                        break;
                    case "name":
                        testObjects = driver.FindElements(By.Name(attribute));
                        break;
                    case "id":
                        try
                        {
                            testObjects = driver.FindElements(By.Id(attribute));
                        }
                        catch (NullReferenceException)
                        {
                            GetElementByAttribute("css", "#" + attribute);
                        }
                        break;
                    case "linktext":
                    case "text":
                    case "link":
                        testObjects = driver.FindElements(By.LinkText(attribute));
                        break;
                    case "xpath":
                        testObjects = driver.FindElements(By.XPath(attribute));
                        break;
                    case "partiallinktext":
                        testObjects = driver.FindElements(By.PartialLinkText(attribute));
                        break;
                    case "tag":
                    case "tagname":
                    case "html tag":
                        testObjects = driver.FindElements(By.TagName(attribute));
                        break;
                    case "class":
                         testObjects = driver.FindElements(By.ClassName(attribute));
                        break;
                    case "classname":
                        testObjects = driver.FindElements(By.ClassName(attribute));
                        break;
                    default:
                        throw new Exception("Locator Type :\"" + attributeType + "\" is undefined");
                }

                if (DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Found " + testObjects.Count().ToString() + " objects matching to locator " + attributeType + "=" + attribute);
                }

                Common.Utility.SetVariable("ElementCount", testObjects.Count().ToString());

                if (testObjects.Count().Equals(0))
                {
                    testObject = null;
                    throw new NoSuchElementException();
                }
                
                if (testObjects.Count().Equals(1))
                {
                    testObject = testObjects.First();
                    return testObject;
                }

                //Variable to store if object is currently displayed
                bool isObjectDisplayed = false;
                firstObject = testObjects.First();

                //For each element, check if it is displayed. If yes, return it
                int counter = 0;
                foreach (IWebElement element in testObjects)
                {
                    counter = counter + 1;

                    //If no displayed object could be located, store first element to test object by default
                    if (!Common.Property.isRecoveryRunning)
                    {
                        testObject = testObjects.First();
                    }

                    try
                    {
                        isObjectDisplayed = element.Displayed;

                        //Print on console about object displayed status
                        if (DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine("Object position: " + counter + ", displayed= " + isObjectDisplayed.ToString());
                        }

                        //Return if object is displayed
                        if (isObjectDisplayed)
                        {
                            this.HightlightObject(element);
                            testObject = element;
                            return testObject;
                        }
                    }
                    catch (Exception displayCheck)
                    {
                        Console.WriteLine("Failed to check display attribute of element at position: " + counter +
                                          "Error:" + displayCheck.Message);
                    }

                }

                //During Browser recovery,we are not supposed to return testobject if it is not displayed on page.
                if (Common.Property.isRecoveryRunning && !isObjectDisplayed)
                {
                    return null;
                }


                // Xpath Generation.
                if (Common.Utility.GetVariable("forcexpath").ToLower().Equals("true"))
                {
                    string xpathString = getEquivalentXpath(testObject);
                    if (xpathString != null)
                    {
                        testObject = null;
                        testObject = driver.FindElement(By.XPath(xpathString));
                        return testObject;
                    }
                }
                return null;
            }
            catch (NoSuchElementException e)
            {
                return null;
            }
            catch (WebDriverException e)
            {
                throw e;
            }
            catch (InvalidOperationException)
            {
                return this.GetSingleElementByAttribute(attributeType, attribute);   
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Create IWebElement using  Element attributes type and Element attribute.
        /// </summary>
        /// <param name="attributeType">string : Type of property</param>
        /// <param name="attribute">string : Property value</param>
        /// <returns>Instance of IwebElement</returns>
        private IWebElement GetSingleElementByAttribute(string attributeType, string attribute)
        {
            try
            {
                switch (attributeType.ToLower())
                {

                    case "cssselector":
                        testObject = driver.FindElement(By.CssSelector(attribute));
                        break;
                    case "css":
                        testObject = driver.FindElement(By.CssSelector(attribute));
                        break;
                    case "name":
                        testObject = driver.FindElement(By.Name(attribute));
                        break;
                    case "id":
                        try
                        {
                            testObject = driver.FindElement(By.Id(attribute));
                        }
                        catch (Exception)
                        {
                            GetSingleElementByAttribute("css", "#" + attribute);
                        }
                        break;
                    case "linktext":
                    case "text":
                    case "link":
                        testObject = driver.FindElement(By.LinkText(attribute));
                        break;
                    case "xpath":
                        testObject = driver.FindElement(By.XPath(attribute));
                        break;
                    case "partiallinktext":
                        testObject = driver.FindElement(By.PartialLinkText(attribute));
                        break;
                    case "tag":
                    case "tagname":
                    case "html tag":
                        testObject = driver.FindElement(By.TagName(attribute));
                        break;
                    case "class":
                         testObject = driver.FindElement(By.ClassName(attribute));
                        break;
                    case "classname":
                        testObject = driver.FindElement(By.ClassName(attribute));
                        break;
                    default:
                        return null;
                }
                //Finding an element from the generated x-path from Iwebelement.
                if (Common.Utility.GetVariable("forcexpath").ToLower().Equals("true"))
                {
                    string xpath = getEquivalentXpath(testObject);
                    if (xpath != null)
                        testObject = driver.FindElement(By.XPath(xpath));
                }

                return testObject;
            }
            catch (NoSuchElementException e)
            {
                Common.KryptonException.writeexception(e);
                return null;
            }
            catch (WebDriverException e)
            {
                Common.KryptonException.writeexception(e);
                throw e;
            }
            catch (Exception e)
            {
                Common.KryptonException.writeexception(e);
                throw e;
            }
        }
        /// <summary>
        /// Get equivalent X-path for IWebElement.
        /// </summary>
        /// <param name="tObject">IWebElement : TestObject</param>
        /// <returns>string : xpath string.</returns>
        private string getEquivalentXpath(IWebElement tObject)
        {
            try
            {
                string scriptResult = (String)((IJavaScriptExecutor)driver).ExecuteScript("gPt=function(c){if(c.id!==''){return'id(\"'+c.id+'\")'} " +
                        "if(c===document.body){return c.tagName}" +
                        "var a=0;var e=c.parentNode.childNodes;" +
                        "for(var b=0;b<e.length;b++){var d=e[b];if(d===c){return gPt(c.parentNode)+'/'+c.tagName+'['+(a+1)+']'}if(d.nodeType===1&&d.tagName===c.tagName){a++}}};" +
                        "return gPt(arguments[0]).toLowerCase();", tObject);
                using (StreamWriter sw = new StreamWriter(@"C:\Krypton\TestLogs"))
                {
                    sw.WriteLine("getequivalantXpath() 5");
                }
                return scriptResult;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        ///Move the mouse to specified IWebElement based on its co-ordinates.
        /// </summary>
        public void mouseMove()
        {
            try
            {
                testObject = this.WaitAndGetElement();
                //Get Position of testObject.
                ICoordinates posTestObject = ((ILocatable)testObject).Coordinates;
                //Get Mouse Control.
                IMouse mouseControl = ((IHasInputDevices)driver).Mouse;
                //Move mouse control to posTestObject.
                mouseControl.MouseMove(posTestObject);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Click on an element using actions object
        /// </summary>
        public void mouseClick()
        {
            try
            {
                addAction("click");
                performAction();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// mouse over on an element using actions object
        /// </summary>
        public void mouseOver()
        {
            try
            {
                addAction("mouseover");
                performAction();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// adds actions to advanced user interaction API
        /// </summary>
        public void addAction(string actionToAdd = "", string data = "")
        {
            try
            {
                switch (actionToAdd.ToLower())
                {
                    case "click":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.Click(testObject);
                        break;
                    case "clickandhold":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.ClickAndHold(testObject);
                        break;
                    case "contextclick":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.ContextClick(testObject);
                        break;
                    case "doubleclick":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.DoubleClick(testObject);
                        break;
                    case "keydown":
                    case "keyup":
                    case "release":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.Release(testObject);
                        break;
                    case "sendkeys":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.SendKeys(testObject, data);
                        break;
                    case "mouseover":
                    case "movetoelement":
                    case "mousemove":
                        testObject = this.WaitAndGetElement();
                        Browser.driverActions.MoveToElement(testObject);
                        break;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Performs actions present in action object
        /// </summary>
        public void performAction()
        {
            try
            {
                //Finally, perform action on webelement
                IAction finalAction = Browser.driverActions.Build();
                finalAction.Perform();
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Method to clear text in text box.
        /// 
        /// </summary>
        public void ClearText()
        {
            try
            {
                testObject = this.WaitAndGetElement();
                testObject.Clear();
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// Method to click on button and Link
        ///
        /// </summary>
        public void Click(Dictionary<int, string> keyWordDic = null, string data = null)
        {
            try
            {
                testObject = this.WaitAndGetElement();               
                DateTime startTime = DateTime.Now;
                if (Common.Utility.GetParameter("runbyevents").Equals("true"))
                {
                    switch (Browser.browserName.ToLower())
                    {
                        case "ie":
                        case "iexplore":
                            this.ExecuteScript(testObject, "arguments[0].click();");
                            try
                            {
                                this.WaitForObjectNotPresent(Common.Utility.GetVariable("ObjectTimeout"),
                                                             Common.Utility.GetVariable("GlobalTimeout"), keyWordDic);
                            }
                            catch
                            {
                                //Nothing to throw in this case
                            }
                            break;
                        default:
                            testObject.Click();
                            break;
                    }
                }
                else
                {
                    try
                    {
                        // Perform Shift+Click only if Shift key is passed as Data for action sheet data in case Chrome.
                        if ("Shift".Equals(data, StringComparison.OrdinalIgnoreCase) && Browser.browserName.Equals(KryptonConstants.BROWSER_CHROME,StringComparison.OrdinalIgnoreCase))
                        {
                            Actions objAction = new Actions(driver);

                            objAction = objAction.KeyDown(Keys.Shift).Click(testObject).KeyUp(Keys.Shift);
                            IAction objIAction = objAction.Build();
                            objAction.Perform();
                        }
                        else
                        {
                            testObject.Click();
                        }

                        // Pause here for .5 sec as on mac safari VerifyTextOnPage after this looks at the page click is on and assumes that page has been received from server
                        // find a better way to do this
                        if (Browser.browserName.Equals("safari"))
                            Thread.Sleep(2000);
                    }
                      catch(ElementNotVisibleException enve)
                      {
                          try { ExecuteScript(testObject, "arguments[0].click();"); }
                          catch { throw enve; }
                        
                       }
                    catch (Exception e)
                     {
                        if (e.Message.ToLower().Contains(Common.exceptions.ERROR_NORESPONSEURL))
                        {
                            testObject.Click();
                        }
                        else
                        {
                           throw e;
                        }                      
                        
                       
                    }
                }

                //measure total time and raise exception if timeout is more than the allowed limit
                DateTime finishTime = DateTime.Now;
                double totalTime = (double)(finishTime - startTime).TotalSeconds;
                foreach (string modifiervalue in keyWordDic.Values)
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
            catch (Exception e)
            {
                if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                {
                    Common.KryptonException.writeexception(e);
                    throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0067").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute).Replace("{ErrorMsg}", e.Message)); // added by 
                }
                else
                    throw;
            }
        }

        public void clickInThread()
        {
            try
            {

                testObject = this.WaitAndGetElement();

                var ts = new CancellationTokenSource();
                CancellationToken ct = ts.Token;
                Task.Factory.StartNew(() =>
                {
                    while (true)
                    {
                        testObject.Click();
                        Thread.Sleep(100);
                        if (ct.IsCancellationRequested)
                        {
                            // another thread decided to cancel
                            Console.WriteLine("task canceled");
                            break;
                        }
                    }
                }, ct);
                // Simulate waiting 10s for the task to complete
                Thread.Sleep(10000);
                // Can't wait anymore => cancel this task
                ts.Cancel();
            }
            catch (Exception e)
            {
                throw e;
            }


        }


        /// <summary>
        ///Method to press non alphabetic key.eg- Enter Key.
        /// </summary>
        /// <param name="key">String : Keyboard key </param>
        public void KeyPress(string key)
        {
            testObject = this.WaitAndGetElement();
            switch (key.ToLower())
            {
                case "arrowdown":
                    testObject.SendKeys(Keys.ArrowDown);
                    break;
                case "enter":
                    testObject.SendKeys(Keys.Enter);
                    break;
                case "add":
                    testObject.SendKeys(Keys.Add);
                    break;
                case "alt":
                    testObject.SendKeys(Keys.Alt);
                    break;
                case "arrowleft":
                    testObject.SendKeys(Keys.ArrowLeft);
                    break;
                case "arrowright":
                    testObject.SendKeys(Keys.ArrowRight);
                    break;
                case "arrowup":
                    testObject.SendKeys(Keys.ArrowUp);
                    break;
                case "backspace":
                    testObject.SendKeys(Keys.Backspace);
                    break;
                case "cancel":
                    testObject.SendKeys(Keys.Cancel);
                    break;
                case "clear":
                    testObject.SendKeys(Keys.Clear);
                    break;
                case "command":
                    testObject.SendKeys(Keys.Command);
                    break;
                case "control":
                case "ctrl":
                    testObject.SendKeys(Keys.Control);
                    break;
                case "decimal":
                    testObject.SendKeys(Keys.Decimal);
                    break;
                case "delete":
                    testObject.SendKeys(Keys.Delete);
                    break;
                case "divide":
                    testObject.SendKeys(Keys.Divide);
                    break;
                case "down":
                    testObject.SendKeys(Keys.Down);
                    break;
                case "end":
                    testObject.SendKeys(Keys.End);
                    break;
                case "equal":
                    testObject.SendKeys(Keys.Equal);
                    break;
                case "escape":
                    testObject.SendKeys(Keys.Escape);
                    break;
                case "f1":
                    testObject.SendKeys(Keys.F1);
                    break;
                case "f10":
                    testObject.SendKeys(Keys.F10);
                    break;
                case "f11":
                    testObject.SendKeys(Keys.F11);
                    break;
                case "f12":
                    testObject.SendKeys(Keys.F12);
                    break;
                case "f2":
                    testObject.SendKeys(Keys.F2);
                    break;
                case "f3":
                    testObject.SendKeys(Keys.F3);
                    break;
                case "f4":
                    testObject.SendKeys(Keys.F4);
                    break;
                case "f5":
                    testObject.SendKeys(Keys.F5);
                    break;
                case "f6":
                    testObject.SendKeys(Keys.F6);
                    break;
                case "f7":
                    testObject.SendKeys(Keys.F7);
                    break;
                case "f8":
                    testObject.SendKeys(Keys.F8);
                    break;
                case "f9":
                    testObject.SendKeys(Keys.F9);
                    break;
                case "help":
                    testObject.SendKeys(Keys.Help);
                    break;
                case "home":
                    testObject.SendKeys(Keys.Home);
                    break;
                case "insert":
                    testObject.SendKeys(Keys.Insert);
                    break;
                case "left":
                    testObject.SendKeys(Keys.Left);
                    break;
                case "leftalt":
                    testObject.SendKeys(Keys.LeftAlt);
                    break;
                case "leftcontrol":
                    testObject.SendKeys(Keys.LeftControl);
                    break;
                case "leftshift":
                    testObject.SendKeys(Keys.LeftShift);
                    break;
                case "meta":
                    testObject.SendKeys(Keys.Meta);
                    break;
                case "multiply":
                    testObject.SendKeys(Keys.Multiply);
                    break;
                case "null":
                    testObject.SendKeys(Keys.Null);
                    break;
                case "numberpad0":
                    testObject.SendKeys(Keys.NumberPad0);
                    break;
                case "numberpad1":
                    testObject.SendKeys(Keys.NumberPad1);
                    break;
                case "numberpad2":
                    testObject.SendKeys(Keys.NumberPad2);
                    break;
                case "numberpad3":
                    testObject.SendKeys(Keys.NumberPad3);
                    break;
                case "numberpad4":
                    testObject.SendKeys(Keys.NumberPad4);
                    break;
                case "numberpad5":
                    testObject.SendKeys(Keys.NumberPad5);
                    break;
                case "numberpad6":
                    testObject.SendKeys(Keys.NumberPad6);
                    break;
                case "numberpad7":
                    testObject.SendKeys(Keys.NumberPad7);
                    break;
                case "numberpad8":
                    testObject.SendKeys(Keys.NumberPad8);
                    break;
                case "numberpad9":
                    testObject.SendKeys(Keys.NumberPad9);
                    break;
                case "pagedown":
                    testObject.SendKeys(Keys.PageDown);
                    break;
                case "pageup":
                    testObject.SendKeys(Keys.PageUp);
                    break;
                case "pause":
                    testObject.SendKeys(Keys.Pause);
                    break;
                case "return":
                    testObject.SendKeys(Keys.Return);
                    break;
                case "right":
                    testObject.SendKeys(Keys.Right);
                    break;
                case "semicolon":
                    testObject.SendKeys(Keys.Semicolon);
                    break;
                case "separator":
                    testObject.SendKeys(Keys.Separator);
                    break;
                case "shift":
                    testObject.SendKeys(Keys.Shift);
                    break;
                case "space":
                    testObject.SendKeys(Keys.Space);
                    break;
                case "subtract":
                    testObject.SendKeys(Keys.Subtract);
                    break;
                case "tab":
                    testObject.SendKeys(Keys.Tab);
                    break;
                case "up":
                    testObject.SendKeys(Keys.Up);
                    break;
                case "ctrl+c":
                    testObject.SendKeys(Keys.Control + "c");
                    break;
                case "ctrl+v":
                    testObject.SendKeys(Keys.Control + "v");
                    break;
                case "ctrl+a":
                    testObject.SendKeys(Keys.Control + "a");
                    break;

            }

        }
        /// <summary>
        /// This will check all specified child checkboxes.
        /// </summary>
        /// <param name="dataContents">String[] : Labels of the checkboxes that have to be checked.</param>
        public void checkMultiple(string[] dataContents)
        {
            try
            {
                //Get the Test Object.
                testObject = WaitAndGetElement();

                //Get All Specified CheckBoxes
                List<IWebElement> AllCheckBoxes = this.getAllCheckBoxes(dataContents);

                //Checking All Specified Checkboxes.
                foreach (IWebElement Checkbox in AllCheckBoxes)
                {
                    if (!Checkbox.Selected)
                    {
                        Checkbox.Click();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// Get All checkbox Objects
        /// </summary>
        /// <param name="dataContents"></param>
        /// <returns></returns>
        private List<IWebElement> getAllCheckBoxes(string[] dataContents)
        {
            try
            {
                //Get All Labels in the test object.
                IList<IWebElement> AllLabels = testObject.FindElements(By.TagName("label"));

                //Declaring List Of all Checkboxes that need to be processed.
                List<IWebElement> AllCheckBoxes = new List<IWebElement>();

                //Processing each Labels one by one.
                foreach (IWebElement Labels in AllLabels)
                {
                    //Checking if Checkbox need to be processed based upon inputs given.
                    if (!checkLabels(dataContents, Labels.Text))
                    {
                        continue; //Continuing to next label.
                    }

                    string forAttribute = string.Empty;

                    IWebElement CheckBoxOfLabel = null;
                    //Fetching 'for' attribute of current label.
                    forAttribute = Labels.GetAttribute("for");
                    //There are checkbox groups in which for is not specified, these groups have different hierarchies.
                    if (!(forAttribute == null))
                    {
                        //Getting checkbox Element.
                        CheckBoxOfLabel = driver.FindElement(By.Id(forAttribute));

                        //Add checkbox to list.
                        AllCheckBoxes.Add(CheckBoxOfLabel);
                    }
                    else //If for attribute is not present in DOM for labels.
                    {
                        //Getting CheckBox Element.
                        CheckBoxOfLabel = Labels.FindElement(By.TagName("input"));

                        //Add checkbox to list.
                        AllCheckBoxes.Add(CheckBoxOfLabel);
                    }
                }
                return AllCheckBoxes;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Uncheck all child checkboxes.
        /// </summary>
        /// <param name="dataContent">string[] : specified the labels of checkboxes.</param>
        public void uncheckMultiple(string[] dataContent)
        {
            try
            {
                //Get Test Object.
                testObject = this.WaitAndGetElement();

                //Get all specified checkboxes.
                List<IWebElement> AllCheckBoxes = this.getAllCheckBoxes(dataContent);

                //UnChecking All Specified Checkboxes.
                foreach (IWebElement Checkbox in AllCheckBoxes)
                {
                    if (Checkbox.Selected)
                    {
                        Checkbox.Click();
                    }
                }

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Check if specified label is the current label.
        /// </summary>
        /// <param name="dataContent"></param>
        /// <param name="LabelText"></param>
        /// <returns></returns>
        private bool checkLabels(string[] dataContent, string LabelText)
        {
            try
            {
                bool checkLabel = false;
                foreach (string data in dataContent)
                {
                    if (data.Trim().Equals(LabelText))
                    {
                        checkLabel = true;
                        break;
                    }
                }
                return checkLabel;
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        /// <summary>
        ///  Method to check on checkbox and radio button.
        /// </summary>
        public void Check(string checkStatus = "")
        {
            try
            {
                testObject = this.WaitAndGetElement();
                bool isSelected = testObject.Selected;

                switch (checkStatus.ToLower().Trim())
                {
                    case "on":
                        if (!isSelected)
                        {
                            testObject.Click();
                        }
                        break;
                    case "off":
                        if (isSelected)
                        {
                            testObject.Click();
                        }
                        break;
                    default:
                        if (!isSelected)
                        {
                            testObject.Click();
                        }
                        break;
                }

            }
            catch (Exception ex)
            {

                if (ex.Message.Contains("Cannot click on element"))
                {
                    try
                    {
                        ExecuteScript(testObject, "arguments[0].click();");  
                    }
                    catch
                    {
                        throw ex;
                    }
                }
                else
                throw;
            }
        }


        /// <summary>
        ///Method to uncheck on check box and radio button.
        /// </summary>
        public void UnCheck()
        {

            try
            {
                this.Check("off");
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// Method to type specied text in text box.
        /// </summary>
        /// <param name="text">string : Text to type on text box object</param>
        public void SendKeys(string text)
        {
            try
            {
              
                    testObject = this.WaitAndGetElement();               
               
                //Check ON or OFF condition for radio button or checkbox. : 
                if (text.Equals("ON", StringComparison.CurrentCultureIgnoreCase) || (text.Equals("OFF", StringComparison.CurrentCultureIgnoreCase)))
                {
                    string objType = string.Empty;
                    objType = testObject.GetAttribute("type");
                    if (objType.Equals("checkbox", StringComparison.CurrentCultureIgnoreCase) ||
                         objType.Equals("radio", StringComparison.CurrentCultureIgnoreCase))
                    {
                        switch (text.ToLower())
                        {
                            case "on":
                                Common.Property.StepDescription = "Check " + testObject.Text;
                                this.Check();
                                break;
                            case "off":
                                Common.Property.StepDescription = "Uncheck " + testObject.Text;
                                this.UnCheck();
                                break;
                        }
                    }
                }
                else
                {
                    //checking for frames.
                    string objTagName = string.Empty;
                    if (objTagName.ToLower().Equals("iframe") || objTagName.ToLower().Equals("frame"))
                    {
                        driver.SwitchTo().Frame(testObject); //Selecting the frame
                        driver.FindElement(By.CssSelector("body")).SendKeys("text"); //entering the data to frame.
                    }
                    else
                    {
                        testObject.Clear();
                        try
                        {
                            testObject.SendKeys(text);
                        }
                        catch(ElementNotVisibleException enve)
                        {
                            try
                            { 
                                ExecuteScript(testObject, string.Format("arguments[0].value='{0}';", text)); 
                            }
                             catch
                            {
                                throw enve;
                            }   
                        }
                        catch (Exception ee)
                        {
                            if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                                throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0069").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // added by 
                            else
                                throw ee;
                        }
                    }
                }
            }
            
            catch (OpenQA.Selenium.StaleElementReferenceException)
            {
                try
                {
                    testObject = this.WaitAndGetElement();
                    testObject.Clear();
                    try
                    {
                        testObject.SendKeys(text);
                    }
                    catch (Exception ee)
                    {
                        if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                            throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0069").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // added by 
                        else
                            throw ee;
                    }
                }
                catch( Exception ex)
                {
                    throw ex;
                }
            }
            catch (Exception e)
            {
                if (e.Message.Contains("Element is no longer attached to the DOM"))
                {
                    testObject = this.WaitAndGetElement();
                    testObject.Clear();
                    try
                    {
                        testObject.SendKeys(text);
                    }
                    catch (Exception ee)
                    {
                        if (objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                            throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0069").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute)); // added by 
                        else
                            throw ee;
                    }
                }
                else
                {
                    throw e;
                }
            }
        }


        /// <summary>
        ///Method to submit form data.
        /// </summary>
        public void Submit()
        {
            try
            {
                testObject = this.WaitAndGetElement();
                testObject.Submit();

            }
            catch (Exception)
            {
                throw;
            }
        }

       
        /// <summary>
        ///Method to fire specified event.
        /// </summary>
        public void FireEvent(string eventName)
        {
            string script = string.Empty;
            string onEventName = string.Empty;

            try
            {
                //Retrieve test object from application
                testObject = this.WaitAndGetElement();

                //Get all events that needs to be fired
                string[] arrEvents = eventName.Split('|');
                string[] arrScripts = new string[arrEvents.Length];


                //Start a for loop for each event, created javascript and store to script array
                for (int i = 0; i < arrEvents.Length; i++)
                {
                    //Retrieve event name from array one by one
                    eventName = arrEvents[i].ToString().Trim().ToLower(); //event must be in lower case :

                    // Replace special characters
                    eventName = Common.Utility.ReplaceSpecialCharactersInString(eventName);

                    //Eventname should start with "on" for internet explorer
                    if (eventName.ToLower().StartsWith("on"))
                    {
                        onEventName = eventName;
                        eventName = eventName.Substring(2);
                    }
                    else
                    {
                        onEventName = "on" + eventName;
                    }

                    //Following script should work good for all browsers
                    script = "var canBubble = false;" + Environment.NewLine +
                            "var element = arguments[0];" + Environment.NewLine +
                            "    if (document.createEventObject()) {" + Environment.NewLine +
                            "        var evt = document.createEventObject();" + Environment.NewLine +
                            "        arguments[0].fireEvent('" + onEventName + "', evt);" + Environment.NewLine +
                            "    }" + Environment.NewLine +
                            "    else {" + Environment.NewLine +
                            "        var evt = document.createEvent(\"HTMLEvents\");" + Environment.NewLine +
                            "        evt.initEvent('" + eventName + "', true, true);" + Environment.NewLine +
                            "        arguments[0].dispatchEvent(evt);" + Environment.NewLine +
                            "    }";


                    //Firefox and others has to be force for this script
                    if (!Browser.browserName.Equals("ie"))
                    {
                        script = "var evt = document.createEvent(\"HTMLEvents\"); evt.initEvent(\"" +
                                 eventName + "\", true, true );return !arguments[0].dispatchEvent(evt);";
                    }

                    //Store scripts in script specific array
                    arrScripts[i] = script;

                }

                //Execute scripts now. This loop needs to be saparate as events needs to be fired as fast as possible
                for (int i = 0; i < arrScripts.Length; i++)
                {
                    this.ExecuteScript(testObject, arrScripts[i]);
                }
            }

            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        ///Method to wait for specified object to become visible.
        /// </summary>
        /// <param name="waitForType">string : Type of condition</param>
        /// <param name="optionDic">Dictionary : Option dictionary</param>
        public string WaitForObject(string waitTime, string globalWaitTime, Dictionary<int, string> optionDic, string modifier)
        {
            try
            {
                int intWaitTime = 0;
                this.modifier = modifier;
                this.optiondic = optionDic;
                if (string.IsNullOrEmpty(globalWaitTime.Trim()))
                {
                    globalWaitTime = "0";
                }
                if (string.IsNullOrEmpty(waitTime))
                {
                    waitTime = globalWaitTime;
                }
                try
                {
                    intWaitTime = Int32.Parse(waitTime);
                }
                catch
                {
                    intWaitTime = Int32.Parse(globalWaitTime);
                }
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(intWaitTime));
                Func<IWebDriver, bool> condition;
                condition = this.VerifyObjectDisplayCondition;
                DateTime dtBefore = System.DateTime.Now;
                wait.Until((Func<IWebDriver, bool>)condition);
                DateTime dtafter = System.DateTime.Now;
                TimeSpan timeSpan = dtafter - dtBefore;
                int timeInterval = timeSpan.Milliseconds;
                return timeInterval.ToString();

            }
            catch (FormatException)
            {
                throw new Exception("GlobalTimeout parameter with value: " + globalWaitTime + "  was not in a correct format");
            
            }
            catch (WebDriverTimeoutException)
            {
                throw new Exception("WebDriverTime Out Exception");
            }

        }

        /// <summary>
        ///Method to wait for specified object to disappear.
        /// </summary>
        /// <param name="waitForType">string : Type of condition</param>
        /// <param name="optionDic">Dictionary : Option dictionary</param>
        public void WaitForObjectNotPresent(string waitTime, string globalWaitTime, Dictionary<int, string> optionDic)
        {
            if (waitTime.Equals(string.Empty))
                {
                    waitTime = globalWaitTime;
                }
                this.optiondic = optionDic;
                
                int intWaitTime = Int32.Parse(waitTime);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(intWaitTime));
                Func<IWebDriver, bool> condition;
                condition = this.VerifyObjectNotDisplayedCondition;
                wait.Until((Func<IWebDriver, bool>)condition);
           
        }
        /// <summary>
        ///Wait for specified test object property to appear.
        /// </summary>
        /// <param name="testData">string : test data</param>
        /// <param name="globalWaitTime">string : Global time out</param>
        /// <param name="optionDic">Dictionary : Option Dictionary</param>
        public void WaitForObjectProperty(string propertyParam, string propertyValueParam, string globalWaitTime, Dictionary<int, string> optionDic)
        {
            try
            {
                property = propertyParam.Trim();
                propertyValue = propertyValueParam.Trim();
                this.optiondic = optionDic;
                int intWaitTime = Int32.Parse(globalWaitTime);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(intWaitTime));
                Func<IWebDriver, bool> condition;
                condition = this.VerifyObjectPropertyCondition;
                wait.Until((Func<IWebDriver, bool>)condition);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        /// <summary>
        ///Method to verify specified test object property.
        /// </summary>
        /// <param name="driver"> selenium driver</param>
        /// <returns>boolean value</returns>
        private bool VerifyObjectPropertyCondition(IWebDriver driver)
        {
            try
            {
                return this.VerifyObjectProperty(property, propertyValue, optiondic);
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (Exception e)
            {
                throw e;
            }


        }

        /// <summary>
        ///Method to verify that object display condition.
        /// </summary>
        /// <param name="driver">Driver instance</param>
        /// <returns>boolean value</returns>
        private bool VerifyObjectDisplayCondition(IWebDriver driver)
        {
            try
            {
                bool status = this.VerifyObjectDisplayed();
                if (!status && modifier.Equals("refresh"))
                {
                    driver.Navigate().Refresh();
                }
                return status;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (StaleElementReferenceException)
            {
                return false;
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        ///Method to verify that object not display condition.
        /// </summary>
        /// <param name="driver">Driver instance</param>
        /// <returns>boolean value</returns>
        private bool VerifyObjectNotDisplayedCondition(IWebDriver driver)
        {
            try
            {
                return !this.VerifyObjectDisplayed();
            }
            catch (NoSuchElementException e)
            {
                return true;
            }
        }

        /// <summary>
        ///Method to perform mouse doubleclick action.
        /// </summary>
        public void DoubleClick()
        {
            try
            {
                selenium = Browser.GetSeleniumOne();
                selenium.DoubleClick(attribute);
            }
            catch (Exception)
            {
                throw;
            }
        }

        
        /// <summary>
        /// Method to select specified item from list.
        /// </summary>
        /// <param name="text">string : option to be select from list</param>
        /// /

        public void SelectItem(string[] itemList, bool selectMultiple = false)
        {
            string AvailbleItemsInList = string.Empty;
            string expectedItems = string.Empty;

            try
            {
                string firstItem = itemList.First();
                foreach (var item in itemList)
                {
                    expectedItems += string.Format("\t" + item);
                }
                testObject = this.WaitAndGetElement();
               if (testObject != null)
                   AvailbleItemsInList = testObject.Text;
                

                #region  KRYPTON0156: Handling conditions where selectItem can be used for radio groups
                if (testObject.GetAttribute("type") != null && testObject.GetAttribute("type").ToLower().Equals("radio"))
                {
                    string radioValue = string.Empty;

                    //Check if given option is also stored in mapping column
                    string valueFromMapping = string.Empty;

                    foreach (KeyValuePair<string, string> mappingKey in dicMapping)
                    {
                        if (mappingKey.Key.ToLower().Equals(firstItem.ToLower()))
                        {
                            valueFromMapping = mappingKey.Value;
                            break;
                        }
                    }

                    foreach (IWebElement radioObject in testObjects)
                    {
                        //Retrieve value property of radio button and compare with passed text
                        string tempRadioValue = radioObject.GetAttribute("value");

                        if (tempRadioValue.ToLower().Equals(firstItem.ToLower()) || tempRadioValue.ToLower().Equals(valueFromMapping.ToLower()))
                        {
                            radioValue = tempRadioValue;
                            testObject = radioObject;
                            break;
                        }
                    }

                    //Determine if a radio option was found or not, and take action based on that
                    if (!radioValue.Equals(string.Empty))
                    {
                        this.ExecuteScript(testObject, "arguments[0].click();");
                        return;
                    }
                    else
                    {
                        Common.Property.Remarks = string.Empty;
                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0022").Replace("{MSG}", firstItem));
                    }
                }
                #endregion

                // Entire section commented above works just fine, but damn slow on IE
                //Following code section is a shortcut for the same, and works reliably and fast on all browsers

                //Clear previus selection if multiple options needs to be selected
                
                    if (selectMultiple)
                    {
                        SelectElement select = new SelectElement(testObject);
                        select.DeselectAll();
                    }
                

                bool IsSelectionSuccessful = false;


                foreach (string optionForSelection in itemList)
                {
                    string optionText = optionForSelection.Trim();


                    IList<IWebElement> allOptions = testObject.FindElements(By.TagName("option"));
                    foreach (IWebElement option in allOptions)
                    {
                        try
                        {
                            string propertyValue = option.GetAttribute("value");
                            if (option.Displayed)
                            {
                                if (option.Text.Contains(optionText))
                                {
                                    IsSelectionSuccessful = true;
                                    option.Click();
                                    break;
                                }
                                else if (propertyValue != null && propertyValue.ToLower().Contains(optionText))
                                {
                                    IsSelectionSuccessful = true;
                                    option.Click();
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex is StaleElementReferenceException)
                            {
                                Thread.Sleep(2500);
                                IsSelectionSuccessful = true;
                                option.Click();
                                break;
                            }
                            else
                                throw new Exception(string.Format("Could not select \"{0}\" in the list with options:\" {1} \". Actual error:{2}", optionText, AvailbleItemsInList, ex.Message));

                        }
                    }

                }
                if (!IsSelectionSuccessful)
                {
                    throw new Exception(string.Format("Could not found \"{0}\" in the list with options:\" {1} \".", expectedItems, AvailbleItemsInList));

                }
            }
            catch
            {
                throw new Exception(string.Format("Could not select \"{0}\" in the list with options:\" {1} \".", expectedItems, AvailbleItemsInList));

            }

        }



        /// <summary>
        /// Method to select option from list by index.
        /// </summary>
        /// <param name="index">string : index of item to be select</param>
        public string SelectItemByIndex(string index)
        {

            try
            {
                int intIndex = Int32.Parse(index) - 1;
                testObject = this.WaitAndGetElement();

                #region  KRYPTON0156: Handling conditions where selectItemByIndex can be used for radio groups
                if (testObject.GetAttribute("type").ToLower().Equals("radio"))
                {
                    testObject = testObjects.ElementAt(intIndex);
                    this.ExecuteScript(testObject, "arguments[0].click();");
                    return testObject.GetAttribute("value");
                }
                #endregion

                SelectElement select = new SelectElement(testObject);
                
                select.SelectByIndex(intIndex);
                IWebElement[] options = select.Options.ToArray();
                return options[intIndex].Text;

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        /// <summary>
        /// Method to return test object property with specified property type.
        /// </summary>
        /// <param name="property">string : Name of the property</param>
        /// <returns>string : Property value</returns>
        public string GetObjectProperty(string property)
        {
            try
            {
                testObject = this.WaitAndGetElement();
                string actualPropertyValue = string.Empty;
                string javascript;

                switch (property.ToLower())
                {
                    case "text":
                        actualPropertyValue = testObject.Text;
                        break;
                    // Edited to work with firefox
                    case "style.background":
                    case "style.backgroundimage":
                    case "style.background-image":
                        switch (Browser.browserName.ToLower())
                        {
                            case "ie":
                                javascript = "return arguments[0].currentStyle.backgroundImage;";
                                break;
                            case "firefox":
                                javascript = "return document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"background-image\");";
                                break;
                            default:
                                javascript = "return arguments[0].currentStyle.backgroundImage;";
                                break;
                        }
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;

                    case "style.color":
                        switch (Browser.browserName.ToLower())
                        {
                            case "ie":
                                javascript = "var color=arguments[0].currentStyle.color;return color;";
                                break;
                            case "firefox":
                                javascript = "var color= document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"color\");var c = color.replace(/rgb\\((.+)\\)/,'$1').replace(/\\s/g,'').split(',');color = '#'+ parseInt(c[0]).toString(16) +''+ parseInt(c[1]).toString(16) +''+ parseInt(c[2]).toString(16);return color;";
                                break;
                            default:
                                javascript = "var color=arguments[0].currentStyle.color;return color;"; ;
                                break;
                        }
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "style.fontweight":
                    case "style.font-weight":
                        switch (Browser.browserName.ToLower())
                        {
                            case "ie":
                                javascript = "var fontWeight= arguments[0].currentStyle.fontWeight;if(fontWeight==700){return 'bold';}if(fontWeight==400){return 'normal';}return fontWait;";
                                break;
                            case "firefox":
                                javascript = "return document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"font-weight\");";
                                break;
                            default:
                                javascript = "return arguments[0].currentStyle.fontWeight;";
                                break;
                        }
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "style.left":
                        switch (Browser.browserName.ToLower())
                        {
                            case "ie":
                                javascript = "return arguments[0].currentStyle.left;";
                                break;
                            case "firefox":
                                javascript = "return document.defaultView.getComputedStyle(arguments[0], '').getPropertyValue(\"left\");";
                                break;
                            default:
                                javascript = "return arguments[0].currentStyle.left;";
                                break;
                        }
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "disabled":
                        javascript = "return arguments[0].disabled;";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;

                    case "tooltip":
                    case "title":
                        javascript = "return title=arguments[0].title;";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "tag":
                    case "html tag":
                    case "htmltag":
                    case "tagname":
                        actualPropertyValue = testObject.GetAttribute("tagName");
                        break;
                    case "scrollheight":
                        javascript = "return arguments[0].scrollHeight;";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "file name":
                    case "filename":
                        javascript = "var src = arguments[0].src; src=src.split(\"/\"); return(src[src.length-1]);";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "height":
                        javascript = "return arguments[0].height;";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "width":
                        javascript = "return arguments[0].width;";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript);
                        break;
                    case "checked":
                        javascript = "return arguments[0].checked;";
                        actualPropertyValue = this.ExecuteScript(testObject, javascript).ToLower();
                        break;
                    case "orientation":
                        // For orientation, calculate width and height
                        int height = Convert.ToInt16(this.GetObjectProperty("height"));
                        int width = Convert.ToInt16(this.GetObjectProperty("width"));

                        //Now, determine if orientation is horizontal, vertical or square
                        actualPropertyValue = "undetermined";
                        if (height > width)
                        {
                            actualPropertyValue = "vertical";
                        }
                        if (height < width)
                        {
                            actualPropertyValue = "horizontal";
                        }
                        if (height == width)
                        {
                            actualPropertyValue = "square";
                        }
                        break;
                    case "selected":
                        // Selected means, either selected option from a radio group, or selected value from list boxes
                        #region KRYPTON0467: Handling conditions where property of selected radio buttons out of a group need to retrieved
                        if (testObject.GetAttribute("type").ToLower().Equals("radio"))
                        {
                            //Out of radio group, retrieve which one is selected
                            string radioValue = string.Empty;
                            javascript = "return arguments[0].checked;";

                            foreach (IWebElement radioObject in testObjects)
                            {
                                radioValue = this.ExecuteScript(radioObject, javascript).ToString();

                                if (radioValue.ToLower().Equals("true") || radioValue.ToLower().Equals("on") || radioValue.ToLower().Equals("1"))
                                {
                                    radioValue = radioObject.GetAttribute("value");
                                    testObject = radioObject;
                                    break;
                                }
                                else
                                {
                                    radioValue = "{no option selected}";
                                }
                            }

                            //Check if given option is also stored in mapping column, if so assign to actual property value
                            actualPropertyValue = radioValue;

                            foreach (KeyValuePair<string, string> mappingKey in dicMapping)
                            {
                                if (mappingKey.Value.ToLower().Equals(radioValue.ToLower()))
                                {
                                    actualPropertyValue = mappingKey.Key;
                                    break;
                                }
                            }
                        }
                        #endregion
                        else
                        {
                            javascript = "return arguments[0].selected;";
                            actualPropertyValue = this.ExecuteScript(testObject, javascript).ToLower();

                        }

                        break;
                    default:
                        actualPropertyValue = testObject.GetAttribute(property);
                        if (actualPropertyValue == null)
                        {
                            throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0023").Replace("{MSG1}", property).Replace("{MSG2}", testObject.Text));
                        }
                        break;
                }
                return actualPropertyValue;
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        ///Method to verify Object visiblity.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyObjectDisplayed()
        {
            try
            {
                IWebElement testObject = (IWebElement)this.GetElement(driver);
                                
                bool status = true;
                if (testObject == null || !testObject.Displayed)
                {
                    Common.Property.Remarks = "Object is not displaying.";
                    status = false;
                }
                return status;
            }
            catch (OpenQA.Selenium.NoSuchElementException nsee)
            {
                throw nsee;
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf(exceptions.ERROR_ELEMENT_NOT_ATTACHED, StringComparison.OrdinalIgnoreCase) >= 0) // done Temp: for Chromedriver issue.
                {
                    throw new Exception("Object is not displaying.");
                }
                if (e.Message.IndexOf("unexpected alert open", StringComparison.OrdinalIgnoreCase) >= 0) // done Temp: for Chromedriver V2.2 issue.
                {
                    throw new Exception("Object is not displaying due to unexpected alert open.");
                }
                else
                    throw e;                
            }
        }


        /// <summary>
        /// Method toverify list item is present.
        /// </summary>
        /// <param name="listItemName">string : option to verify</param>
        /// <returns>boolean value</returns>
        public bool VerifyListItemPresent(string listItemName)
        {
            try
            {
                testObject = this.WaitAndGetElement();
                SelectElement select = new SelectElement(testObject);
                IList<IWebElement> option = select.Options;
                bool flag = false;
                foreach (IWebElement element in option)
                {
                    if (element.Text.Equals(listItemName))
                    {
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Common.Property.Remarks = "List Item \"" + listItemName + "\" does not match with \"" + testObject.Text + "\" values in list";
                }
                return flag;
            }
            catch (Exception)
            {
                throw;
            }
        }


        

        /// <summary>
        /// Method to verify list itme is not present in list.
        /// </summary>
        /// <param name="listItemName">string : option to verify</param>
        /// <returns>boolean value</returns>
        public bool VerifyListItemNotPresent(string listItemName)
        {
            try
            {
                bool status = this.VerifyListItemPresent(listItemName);
                Common.Property.Remarks = string.Empty;
                if (status)
                {
                    Common.Property.Remarks = "List Item \"" + listItemName + "\" is present";
                }
                return !status;
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        /// Method to verify object not present in web page.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyObjectPresent()
        {
            try
            {
                testObject = this.WaitAndGetElement();
                if (testObject != null)
                {
                    IWebElement element = (IWebElement)testObject;
                    bool status = element.Displayed;
                    if (!status)
                    {
                        Common.Property.Remarks = "Element is present but not displayed.";
                    }
                    return status;

                }
                else
                {
                    Common.Property.Remarks = "Element is not present.";
                    return false;
                }

               
               

            }
            catch (NoSuchElementException e)
            {
                Common.Property.Remarks = e.Message;
                return false;
            }
            catch (Exception)
            {
                throw;
            }

        }



        /// <summary>
        ///Method to verify test object is present.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyObjectNotPresent()
        {

            try
            {
                return this.WaitNonPresenceAndGetElement();
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        ///Method to verify object property with specified property.
        /// </summary>
        /// <param name="property">string : Property name</param>
        /// <returns>boolean value</returns>
        public bool VerifyObjectProperty(string property, string propertyValue, Dictionary<int, string> KeywordDic)
        {
            try

            {
                bool isKeyVerified;
                string actualPropertyValue = this.GetObjectProperty(property);
                if (!KeywordDic.Count.Equals(0))
                {
                    isKeyVerified = Common.Utility.doKeywordMatch(propertyValue, actualPropertyValue);
                }
                else
                {
                    isKeyVerified = actualPropertyValue.Equals(propertyValue);
                }
                if (!isKeyVerified)
                {
                    Common.Property.Remarks = "Property -\"" + property + "\", actual value \"" + actualPropertyValue + "\" does not match with expected value - \"" + propertyValue + "\".";
                }
                else
                    Common.Property.Remarks = "Property -\"" + property + "\", actual value \"" + actualPropertyValue + "\" match with expected value - \"" + propertyValue + "\".";

                return isKeyVerified;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public bool VerifyObjectPropertyNot(string property, string propertyValue, Dictionary<int, string> KeywordDic)
        {
            bool status = this.VerifyObjectProperty(property, propertyValue, KeywordDic);
            Common.Property.Remarks = string.Empty;
            if (status)
            {
                Common.Property.Remarks = "Property \"" + property + "\" matches with actual value \"" + propertyValue;
            }
            return !status;
        }

        /// <summary>
        ///Method to hightlight current test object in web page.
        /// </summary>
        private void HightlightObject(IWebElement tObject = null)
        {
            
                if (DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase)
                    && Browser.isRemoteExecution.ToString().Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    string propertyName = "outline";
                    string originalColor = "none";
                    string highlightColor = "#00ff00 solid 3px";

                    try
                    {
                        //This works with internet explorer
                        originalColor = this.ExecuteScript(tObject, "return arguments[0].currentStyle." + propertyName);
                    }
                    catch
                    {
                        //This works with firefox, chrome and possibly others
                        originalColor = this.ExecuteScript(tObject, "return arguments[0].style." + propertyName);
                    }

                    this.ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + highlightColor + "'");
                    Thread.Sleep(50);
                    this.ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + originalColor + "'");
                    Thread.Sleep(50);
                    this.ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + highlightColor + "'");
                    Thread.Sleep(50);
                    this.ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + originalColor + "'");
                    //Thread.Sleep(10);
                }

            
        }


        /// <summary>
        /// This will attempt to fire javascript on an object
        ///  // Added code for Safari Driver
        /// </summary>
        private string ExecuteScript(IWebElement tObject = null, string scriptToExecute = "")
        {
           
                try
                {
                    object scriptRet = null;                  
                    if (Browser.isRemoteExecution.Equals("false"))
                    {
                        switch (Browser.browserName.ToLower())
                        {
                            case KryptonConstants.BROWSER_FIREFOX:
                                FirefoxDriver ffscriptdriver = (FirefoxDriver)driver;
                                scriptRet = ffscriptdriver.ExecuteScript(scriptToExecute, tObject);
                                if (scriptRet != null)
                                {
                                    return scriptRet.ToString();
                                }
                                break;
                            case KryptonConstants.BROWSER_IE:                                
                                InternetExplorerDriver iescriptdriver = (InternetExplorerDriver)driver;
                        
                                scriptRet = iescriptdriver.ExecuteScript(scriptToExecute, tObject);
                                if (scriptRet != null)
                                {
                                    return scriptRet.ToString();
                                }
                                break;
                            case KryptonConstants.BROWSER_CHROME:
                                ChromeDriver crscriptdriver = (ChromeDriver)driver;
                                scriptRet = crscriptdriver.ExecuteScript(scriptToExecute, tObject);
                                if (scriptRet != null)
                                {
                                    return scriptRet.ToString();
                                }
                                break;
                            case KryptonConstants.BROWSER_SAFARI:
                                SafariDriver safacriptdriver = (SafariDriver)driver;
                                scriptRet = safacriptdriver.ExecuteScript(scriptToExecute, tObject);
                                if (scriptRet != null)
                                {
                                    return scriptRet.ToString();
                                }
                                break;

                            default:
                                return string.Empty;
                        }
                    }
                    else
                    {
                    // Create RemoteWebDriver reference variable in order to execute java script.
                        RemoteWebDriver rdscriptriver = (RemoteWebDriver)driver;
                        scriptRet = rdscriptriver.ExecuteScript(scriptToExecute, tObject);
                        if (scriptRet != null)
                        {
                            return scriptRet.ToString();
                        }
                    }
                    return string.Empty;//empty string is return if script return null value.
                }    
                catch (Exception e)
                {
                    if (testObject != null && objDataRow.Count>0)
                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0068").Replace("{MSG3}", objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", objDataRow["parent"]).Replace("{MSG1}", attributeType).Replace("{MSG2}", attribute));
                    else
                        throw new Exception("Java Script Error :  "+ e.Message);
                }
            }
                     
        /// <summary>
        ///Method to exceute javascript on web page or test object.
        /// </summary>
        /// <param name="scriptToExecute">String : Javascript to execute</param>
        /// <returns>string : return value from javascript</returns>
        public string ExecuteStatement(string scriptToExecute)
        {
            try
            {
                if (attributeType == null || attributeType.Equals(string.Empty))
                {
                    testObject = null;
                }
                else
                {
                    testObject = this.GetElement(driver);
                    
                }               
              return this.ExecuteScript(testObject, scriptToExecute);                
            }
            catch (Exception e)
            {
                Browser.switchToMostRecentBrowser();
                return this.ExecuteScript(testObject, scriptToExecute);
            }
        }

        /// <summary>
        ///Method to set property present in DOM
        /// </summary>
        /// <param name="property">string : Name of property</param>
        /// <param name="propertyValue">string : Value of property</param>
        public void SetAttribute(string property, string propertyValue)
        {
            try
            {
                testObject = this.WaitAndGetElement();
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
                this.ExecuteScript(testObject, javascript);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Method to drag the source object on the target object.
        /// </summary>
        /// <param name="targetObjDic">Dictionary : Target object details</param>
        public void DragAndDrop(Dictionary<string, string> targetObjDic)
        {
            try
            {
                IWebElement sourceObject = this.WaitAndGetElement();

                // Set second object's information in dictionary
                this.SetObjDataRow(targetObjDic);
                IWebElement targetObject = this.WaitAndGetElement();

                Actions action = new Actions(driver);
                action.DragAndDrop(sourceObject, targetObject).Perform();
                Thread.Sleep(100);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        
        /// <summary>
        /// Upload the file
        /// </summary>
        /// <param name="path"></param>
        public bool UploadFile(string path , string element)
        {
            string p = @"" + path;
            if (!File.Exists(p))
            {
                return false;
            }
            InternetExplorerDriver iedriver = (InternetExplorerDriver)driver;
            iedriver.FindElement(By.Id(element)).SendKeys(p);
            return true;
        }
    }
}
