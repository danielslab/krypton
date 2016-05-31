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
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Safari;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using Common;
using System.IO;
using Driver.Browsers;


namespace Driver
{
    public class TestObject : ITestObject
    {
        public static IWebDriver Driver;
        private IWebElement _testObject;
        private IWebElement _firstObject;    //This will contain first object from collection
        private System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> _testObjects;
        public static string AttributeType;
        public static string Attribute;
        private string _property;
        private string _propertyValue;
        private Dictionary<string, string> _objDataRow = new Dictionary<string, string>();
        public IWebElement FrameObject;
        //Dictionary to store values from mapping column of object repository
        private readonly Dictionary<string, string> _dicMapping = new Dictionary<string, string>();
        private Dictionary<int, string> _optiondic;
        private Selenium.DefaultSelenium _selenium;
        public string DebugMode = Property.DebugMode;
        public static WebDriverWait ObjloadingWait;
        private readonly List<string> _stepsForIgnoreWait = new List<string> { "settestmode", "verifyobjectnotpresent", "verifyobjectnotdisplayed" };
        public Func<IWebDriver, bool> DelIsObjectLoaded;
        private string _modifier = string.Empty;

        public TestObject(string waitTimeForObject = null)
        {
            Driver = Browser.Driver;
            if (waitTimeForObject != null && Driver != null)
            {
                ObjloadingWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(Double.Parse(waitTimeForObject)));
            }
        }


        /// <summary>
        /// set test object information dictionary.
        /// </summary>
        /// <param name="objDataRow">Dictionary containing test object information</param>
        /// <param name="currentStepAction"></param>
        public void SetObjDataRow(Dictionary<string, string> objDataRow, string currentStepAction = "")
        {

            _dicMapping.Clear();
            if (objDataRow.Count > 0)
            {
                _objDataRow = objDataRow;
                AttributeType = GetData(KryptonConstants.HOW);
                Attribute = GetData(KryptonConstants.WHAT);

                //Retrive mapping column from object repository and parse
                string mapping = objDataRow[KryptonConstants.MAPPING];
                if (!mapping.Equals(string.Empty))
                {
                    Array arrMapping = mapping.Split('|');
                    foreach (string mappingPair in arrMapping)
                    {
                        var mapName = mappingPair.Split('=')[0].Trim();
                        var mapValue = mappingPair.Split('=')[1].Trim();
                        _dicMapping.Add(mapName, mapValue);
                    }
                }

            }
            else
            {
                _objDataRow = objDataRow;
                AttributeType = string.Empty;
                Attribute = string.Empty;
            }
            RecoveryScenarios.CacheAttribute(AttributeType, Attribute);
            //wait for object
            By obBy = ExSelenium.GetSelectionMethod(AttributeType, Attribute);
            if (obBy != null)
            {
                ExSelenium.WaitForElement(obBy, _stepsForIgnoreWait.Contains(currentStepAction.ToLower()) ? 02 : 10,
                    true);
            }
        }

        //When Object Definition is embedded in the test case sheet instead of OR.
        public void SetObjDataRow(String testObject)
        {
            if (testObject != null && testObject.Split(':').Length >= 2)
            {
                AttributeType = testObject.Split(':')[0].Trim();
                Attribute = testObject.Split(':')[1].Trim();
            }
            else
            {
                AttributeType = string.Empty;
                Attribute = string.Empty;
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
                        return _objDataRow[KryptonConstants.PARENT];
                    case "child":
                        return _objDataRow[KryptonConstants.TEST_OBJECT];
                    case KryptonConstants.LOGICAL_NAME:
                        return _objDataRow[KryptonConstants.LOGICAL_NAME];
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
            catch (KeyNotFoundException)
            {
                throw new KeyNotFoundException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0008"));
            }
        }


        /// <summary>
        /// Funtion to wait until the webpage fully loaded.
        /// </summary>
        /// <returns>timeOut : Time until web page waits.</returns>
        public void WaitForPage(string timeOut)
        {
            try
            {
                double time = Convert.ToDouble(timeOut);
                IWait<IWebDriver> wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(time));
                wait.Until(driver1 => ((IJavaScriptExecutor)Driver).ExecuteScript("return document.readyState").Equals("complete"));
            }
            catch (TimeoutException exeTimeOut)
            {
                KryptonException.Writeexception(exeTimeOut);
                Console.WriteLine(ConsoleMessages.WEB_PAGE_FAILE_LOAD, exeTimeOut);
            }
            catch (Exception ex)
            {
                KryptonException.Writeexception(ex);
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
                if (Property.IsRecoveryRunning || Property.NoWait)
                {
                    GetElement(Driver);
                }
                else
                {
                    Func<IWebDriver, IWebElement> delIsObjectLoaded = GetElement;
                    ObjloadingWait.Until(delIsObjectLoaded);
                }

                //Return first object if no object could be located, and no visible objects were still present
                if ((_testObjects.Any())
                        && !Property.IsRecoveryRunning
                        && _testObject == null)
                {
                    _testObject = _testObjects.First();
                }
                return _testObject;

            }
            catch (Exception e)
            {
                if (_objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute));
                if (Attribute.Equals(String.Empty) && AttributeType.Equals(String.Empty))
                    throw new Exception("Could not found object :  " + e.Message);
                throw new NoSuchElementException(Exceptions.ERROR_OBJECTNOTFOUND + " { method: " + AttributeType + ",  selector: " + Attribute + "  }");
            }
        }

        /// <summary>
        /// Function verify the sort order of the contents in specified test element.
        /// </summary>
        /// <param name="propertyForSorting">string : Property for extacting element's content.</param>
        /// <param name="sortOrder">string : Sort Order</param>
        public bool VerifySortOrder(string propertyForSorting, string sortOrder = "")
        {
            char splitCriteria = '\n';
            string[] lstOfContents = null;
            _testObject = WaitAndGetElement();
            switch (propertyForSorting.ToLower())
            {
                case "text":
                    var textContent = _testObject.Text;
                    lstOfContents = textContent.Split(splitCriteria);
                    break;
            }

            var result = isSorted(lstOfContents, sortOrder);
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
                        if (String.Compare(strArray[i], strArray[i + 1], StringComparison.Ordinal) < 0)
                        {
                            return false;
                        }
                    }
                    return true;
                }
                for (int i = 1; i < strArray.Length; i++)
                {
                    if (String.Compare(strArray[i - 1], strArray[i], StringComparison.Ordinal) > 0)
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (NullReferenceException)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0056"));
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
                DelIsObjectLoaded = VerifyObjectNotDisplayedCondition;
                ObjloadingWait.Until(DelIsObjectLoaded);
                return true;
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf(Exceptions.ERROR_OBJECTREPOSITORY, StringComparison.OrdinalIgnoreCase) >= 0)
                    throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0051").Replace("{MSG}", e.Message));
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0020"));
            }
        }


        ///  <summary>
        /// Search frames in webpage if found search testobject in searched frames.If no frame found search testobject direct into 
        ///  </summary>
        /// <returns>IwebElement Instance</returns>
        private IWebElement GetElement(IWebDriver driver)
        {
            try
            {
                #region search element direct to web page.
                if (Attribute.Equals(String.Empty) && AttributeType.Equals(String.Empty))
                    throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0021"));

                //Make a default attempt, but it might throw errors in case previous action have closed active browser
                var element = GetElementByAttribute(AttributeType, Attribute);
                if (element != null)
                {
                    return element;
                }
                #endregion
                #region switch to most recent browser and handle IE certificate warning and try again
                try
                {
                    global::Driver.Browsers.Browser.SwitchToMostRecentBrowser();
                    element = GetElementByAttribute(AttributeType, Attribute);
                    if (element != null)
                    {
                        return element;
                    }
                }
                catch (Exception e)
                {
                    KryptonException.Writeexception(e);
                }
                #endregion

                #region recover browser and try again
                //Temporarily commented as it is creating wrong attributes
                if (!Property.IsRecoveryRunning)
                    RecoveryScenarios.RecoverFromBrowsers();
                element = GetElementByAttribute(AttributeType, Attribute);
                if (element != null)
                {
                    return element;
                }
                #endregion

                #region search element in frames
                string frame = GetData(KryptonConstants.MAPPING);

                if (!frame.Equals(string.Empty))
                {
                    string[] keywords = frame.Split('=');
                    frame = keywords[0].ToLower().Trim().Equals("frame") ? keywords[1].Trim() : string.Empty;
                }

                switch (frame.Equals(string.Empty))
                {
                    case true:
                        // Handles sub frames.
                        FrameObject = null;
                        GetElementFromFrames();
                        if (FrameObject != null)
                        {
                            return FrameObject;
                        }
                        throw new NoSuchElementException();
                    case false:
                        driver.SwitchTo().Frame(frame);
                        FrameObject = GetElementByAttribute(AttributeType, Attribute);
                        if (FrameObject != null)
                        {
                            return FrameObject;
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
                if (!Property.IsRecoveryRunning && !Property.NoWait)
                {
                    RecoveryScenarios.RecoverFromBrowsers();
                    IWebElement element = GetElementByAttribute(AttributeType, Attribute);

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
            catch (StaleElementReferenceException)
            {
                return null;
            }
            catch (WebDriverException)
            {
                throw;
            }
            catch (Exception)
            {
                if (_objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                    throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0019").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute)); // updated By 
                throw;
            }
        }

        /// <summary>
        ///Get Frames collection in which the driver is in focus.
        /// </summary>
        /// <returns>IEnumerable<IWebElement/> collection</returns>
        private IEnumerable<IWebElement> getFrameElement()
        {

            IEnumerable<IWebElement> frames = Driver.FindElements(By.TagName("frame"));
            IEnumerable<IWebElement> iframes = Driver.FindElements(By.TagName("iframe"));
            IEnumerable<IWebElement> openiframe = Driver.FindElements(By.CssSelector("#open_iframe"));
            IEnumerable<IWebElement> elements = openiframe.Union(frames.Union(iframes));
            return elements;
        }

        /// <summary>
        ///Iterate through all sub frames to get testObject.
        /// </summary>
        /// <returns>IWebElement : TestObject found.</returns>
        private void GetElementFromFrames()
        {
            IEnumerable<IWebElement> elements;
            try
            {
                elements = getFrameElement();
            }
            catch (InvalidOperationException)
            {
                throw new NoSuchElementException();
            }
            if (elements != null && elements.Count() > 0)
            {
                foreach (IWebElement element1 in elements)
                {
                    if (FrameObject != null)
                    {
                        break;
                    }
                    IWebElement testobject;
                    try
                    {
                        Driver.SwitchTo().Frame(element1);

                        testobject = GetElementByAttribute(AttributeType, Attribute);

                        if (testobject != null)
                        {
                            FrameObject = testobject;
                            break;
                        }
                    }
                    catch
                    {
                        // Sometime, when switching to a frame, staleElement exception occures
                        //so switching to default contents becomes necessary before switching to frame/ iFrame in question
                        try
                        {
                            Driver.SwitchTo().DefaultContent();
                            Thread.Sleep(100);
                            Driver.SwitchTo().Frame(element1);
                            testobject = GetElementByAttribute(AttributeType, Attribute);

                            if (testobject != null)
                            {
                                FrameObject = testobject;
                                break;
                            }
                            GetElementFromFrames();
                        }
                        catch (Exception)
                        {
                            // ignored
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
                        _testObjects = Driver.FindElements(By.CssSelector(attribute));
                        break;
                    case "css":
                        _testObjects = Driver.FindElements(By.CssSelector(attribute));
                        break;
                    case "name":
                        _testObjects = Driver.FindElements(By.Name(attribute));
                        break;
                    case "id":
                        try
                        {
                            _testObjects = Driver.FindElements(By.Id(attribute));
                        }
                        catch (NullReferenceException)
                        {
                            GetElementByAttribute("css", "#" + attribute);
                        }
                        break;
                    case "linktext":
                    case "text":
                    case "link":
                        _testObjects = Driver.FindElements(By.LinkText(attribute));
                        break;
                    case "xpath":
                        _testObjects = Driver.FindElements(By.XPath(attribute));
                        break;
                    case "partiallinktext":
                        _testObjects = Driver.FindElements(By.PartialLinkText(attribute));
                        break;
                    case "tag":
                    case "tagname":
                    case "html tag":
                        _testObjects = Driver.FindElements(By.TagName(attribute));
                        break;
                    case "class":
                        _testObjects = Driver.FindElements(By.ClassName(attribute));
                        break;
                    case "classname":
                        _testObjects = Driver.FindElements(By.ClassName(attribute));
                        break;
                    default:
                        throw new Exception("Locator Type :\"" + attributeType + "\" is undefined");
                }

                if (DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Found " + _testObjects.Count + " objects matching to locator " + attributeType + "=" + attribute);
                }

                Utility.SetVariable("ElementCount", _testObjects.Count.ToString());

                if (_testObjects.Count.Equals(0))
                {
                    _testObject = null;
                    throw new NoSuchElementException();
                }

                if (_testObjects.Count.Equals(1))
                {
                    _testObject = _testObjects.First();
                    return _testObject;
                }

                //Variable to store if object is currently displayed
                bool isObjectDisplayed = false;
                _firstObject = _testObjects.First();

                //For each element, check if it is displayed. If yes, return it
                int counter = 0;
                foreach (IWebElement element in _testObjects)
                {
                    counter = counter + 1;

                    //If no displayed object could be located, store first element to test object by default
                    if (!Property.IsRecoveryRunning)
                    {
                        _testObject = _testObjects.First();
                    }

                    try
                    {
                        isObjectDisplayed = element.Displayed;

                        //Print on console about object displayed status
                        if (DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine("Object position: " + counter + ", displayed= " + isObjectDisplayed);
                        }

                        //Return if object is displayed
                        if (isObjectDisplayed)
                        {
                            HightlightObject(element);
                            _testObject = element;
                            return _testObject;
                        }
                    }
                    catch (Exception displayCheck)
                    {
                        Console.WriteLine("Failed to check display attribute of element at position: " + counter +
                                          "Error:" + displayCheck.Message);
                    }

                }

                //During Browser recovery,we are not supposed to return testobject if it is not displayed on page.
                if (Property.IsRecoveryRunning && !isObjectDisplayed)
                {
                    return null;
                }


                // Xpath Generation.
                if (Utility.GetVariable("forcexpath").ToLower().Equals("true"))
                {
                    string xpathString = getEquivalentXpath(_testObject);
                    if (xpathString != null)
                    {
                        _testObject = null;
                        _testObject = Driver.FindElement(By.XPath(xpathString));
                        return _testObject;
                    }
                }
                return null;
            }
            catch (NoSuchElementException)
            {
                return null;
            }
            catch (InvalidOperationException)
            {
                return GetSingleElementByAttribute(attributeType, attribute);
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
                        _testObject = Driver.FindElement(By.CssSelector(attribute));
                        break;
                    case "css":
                        _testObject = Driver.FindElement(By.CssSelector(attribute));
                        break;
                    case "name":
                        _testObject = Driver.FindElement(By.Name(attribute));
                        break;
                    case "id":
                        try
                        {
                            _testObject = Driver.FindElement(By.Id(attribute));
                        }
                        catch (Exception)
                        {
                            GetSingleElementByAttribute("css", "#" + attribute);
                        }
                        break;
                    case "linktext":
                    case "text":
                    case "link":
                        _testObject = Driver.FindElement(By.LinkText(attribute));
                        break;
                    case "xpath":
                        _testObject = Driver.FindElement(By.XPath(attribute));
                        break;
                    case "partiallinktext":
                        _testObject = Driver.FindElement(By.PartialLinkText(attribute));
                        break;
                    case "tag":
                    case "tagname":
                    case "html tag":
                        _testObject = Driver.FindElement(By.TagName(attribute));
                        break;
                    case "class":
                        _testObject = Driver.FindElement(By.ClassName(attribute));
                        break;
                    case "classname":
                        _testObject = Driver.FindElement(By.ClassName(attribute));
                        break;
                    default:
                        return null;
                }
                //Finding an element from the generated x-path from Iwebelement.
                if (Utility.GetVariable("forcexpath").ToLower().Equals("true"))
                {
                    string xpath = getEquivalentXpath(_testObject);
                    if (xpath != null)
                        _testObject = Driver.FindElement(By.XPath(xpath));
                }

                return _testObject;
            }
            catch (NoSuchElementException e)
            {
                KryptonException.Writeexception(e);
                return null;
            }
            catch (WebDriverException e)
            {
                KryptonException.Writeexception(e);
                throw;
            }
            catch (Exception e)
            {
                KryptonException.Writeexception(e);
                throw;
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
                string scriptResult = (String)((IJavaScriptExecutor)Driver).ExecuteScript("gPt=function(c){if(c.id!==''){return'id(\"'+c.id+'\")'} " +
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
        public void MouseMove()
        {
            _testObject = WaitAndGetElement();
            //Get Position of testObject.
            ICoordinates posTestObject = ((ILocatable)_testObject).Coordinates;
            //Get Mouse Control.
            IMouse mouseControl = ((IHasInputDevices)Driver).Mouse;
            //Move mouse control to posTestObject.
            mouseControl.MouseMove(posTestObject);
        }

        /// <summary>
        /// Click on an element using actions object
        /// </summary>
        public void MouseClick()
        {
            AddAction("click");
            PerformAction();
        }

        /// <summary>
        /// mouse over on an element using actions object
        /// </summary>
        public void MouseOver()
        {
            AddAction("mouseover");
            PerformAction();
        }

        /// <summary>
        /// adds actions to advanced user interaction API
        /// </summary>
        public void AddAction(string actionToAdd = "", string data = "")
        {
            switch (actionToAdd.ToLower())
            {
                case "click":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.Click(_testObject);
                    break;
                case "clickandhold":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.ClickAndHold(_testObject);
                    break;
                case "contextclick":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.ContextClick(_testObject);
                    break;
                case "doubleclick":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.DoubleClick(_testObject);
                    break;
                case "keydown":
                case "keyup":
                case "release":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.Release(_testObject);
                    break;
                case "sendkeys":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.SendKeys(_testObject, data);
                    break;
                case "mouseover":
                case "movetoelement":
                case "mousemove":
                    _testObject = WaitAndGetElement();
                    Browser.DriverActions.MoveToElement(_testObject);
                    break;
            }
        }

        /// <summary>
        /// Performs actions present in action object
        /// </summary>
        public void PerformAction()
        {
            //Finally, perform action on webelement
            IAction finalAction = Browser.DriverActions.Build();
            finalAction.Perform();
        }

        /// <summary>
        /// Method to clear text in text box.
        /// 
        /// </summary>
        public void ClearText()
        {
            _testObject = WaitAndGetElement();
            _testObject.Clear();
        }


        /// <summary>
        /// Method to click on button and Link
        ///
        /// </summary>
        public void Click(Dictionary<int, string> keyWordDic = null, string data = null)
        {
            try
            {
                _testObject = WaitAndGetElement();
                DateTime startTime = DateTime.Now;
                if (Utility.GetParameter("runbyevents").Equals("true"))
                {
                    switch (Browser.BrowserName.ToLower())
                    {
                        case "ie":
                        case "iexplore":
                            ExecuteScript(_testObject, "arguments[0].click();");
                            try
                            {
                                WaitForObjectNotPresent(Utility.GetVariable("ObjectTimeout"),
                                                             Utility.GetVariable("GlobalTimeout"), keyWordDic);
                            }
                            catch
                            {
                                //Nothing to throw in this case
                            }
                            break;
                        default:
                            _testObject.Click();
                            break;
                    }
                }
                else
                {
                    try
                    {
                        // Perform Shift+Click only if Shift key is passed as Data for action sheet data in case Chrome.
                        if ("Shift".Equals(data, StringComparison.OrdinalIgnoreCase) && Browser.BrowserName.Equals(KryptonConstants.BROWSER_CHROME, StringComparison.OrdinalIgnoreCase))
                        {
                            Actions objAction = new Actions(Driver);

                            objAction = objAction.KeyDown(Keys.Shift).Click(_testObject).KeyUp(Keys.Shift);
                            objAction.Build();
                            objAction.Perform();
                        }
                        else
                        {
                            _testObject.Click();
                        }

                        // Pause here for .5 sec as on mac safari VerifyTextOnPage after this looks at the page click is on and assumes that page has been received from server
                        // find a better way to do this
                        if (Browser.BrowserName.Equals("safari"))
                            Thread.Sleep(2000);
                    }
                    catch (ElementNotVisibleException enve)
                    {
                        try { ExecuteScript(_testObject, "arguments[0].click();"); }
                        catch { throw enve; }

                    }
                    catch (Exception e)
                    {
                        if (e.Message.ToLower().Contains(Exceptions.ERROR_NORESPONSEURL))
                        {
                            _testObject.Click();
                        }
                        else
                        {
                            throw;
                        }
                    }
                }

                //measure total time and raise exception if timeout is more than the allowed limit
                DateTime finishTime = DateTime.Now;
                double totalTime = (finishTime - startTime).TotalSeconds;
                if (keyWordDic != null)
                    foreach (string modifiervalue in keyWordDic.Values)
                    {
                        if (modifiervalue.ToLower().Contains("timeout="))
                        {
                            double timeout = double.Parse(modifiervalue.Split('=').Last());
                            if (totalTime > timeout)
                            {
                                throw new Exception("Page load took " + totalTime.ToString(CultureInfo.InvariantCulture) + " seconds to load against expected time of " + timeout + " seconds.");
                            }
                            Property.Remarks = "Page load took " + totalTime.ToString(CultureInfo.InvariantCulture) + " seconds to load against expected time of " + timeout + " seconds.";
                        }
                    }
            }
            catch (Exception e)
            {
                if (_objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                {
                    KryptonException.Writeexception(e);
                    throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0067").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute).Replace("{ErrorMsg}", e.Message)); // added by 
                }
                throw;
            }
        }

        public void ClickInThread()
        {
            _testObject = WaitAndGetElement();
            var ts = new CancellationTokenSource();
            CancellationToken ct = ts.Token;
            Task.Factory.StartNew(() =>
            {
                while (true)
                {
                    _testObject.Click();
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


        /// <summary>
        ///Method to press non alphabetic key.eg- Enter Key.
        /// </summary>
        /// <param name="key">String : Keyboard key </param>
        public void KeyPress(string key)
        {
            _testObject = WaitAndGetElement();
            switch (key.ToLower())
            {
                case "arrowdown":
                    _testObject.SendKeys(Keys.ArrowDown);
                    break;
                case "enter":
                    _testObject.SendKeys(Keys.Enter);
                    break;
                case "add":
                    _testObject.SendKeys(Keys.Add);
                    break;
                case "alt":
                    _testObject.SendKeys(Keys.Alt);
                    break;
                case "arrowleft":
                    _testObject.SendKeys(Keys.ArrowLeft);
                    break;
                case "arrowright":
                    _testObject.SendKeys(Keys.ArrowRight);
                    break;
                case "arrowup":
                    _testObject.SendKeys(Keys.ArrowUp);
                    break;
                case "backspace":
                    _testObject.SendKeys(Keys.Backspace);
                    break;
                case "cancel":
                    _testObject.SendKeys(Keys.Cancel);
                    break;
                case "clear":
                    _testObject.SendKeys(Keys.Clear);
                    break;
                case "command":
                    _testObject.SendKeys(Keys.Command);
                    break;
                case "control":
                case "ctrl":
                    _testObject.SendKeys(Keys.Control);
                    break;
                case "decimal":
                    _testObject.SendKeys(Keys.Decimal);
                    break;
                case "delete":
                    _testObject.SendKeys(Keys.Delete);
                    break;
                case "divide":
                    _testObject.SendKeys(Keys.Divide);
                    break;
                case "down":
                    _testObject.SendKeys(Keys.Down);
                    break;
                case "end":
                    _testObject.SendKeys(Keys.End);
                    break;
                case "equal":
                    _testObject.SendKeys(Keys.Equal);
                    break;
                case "escape":
                    _testObject.SendKeys(Keys.Escape);
                    break;
                case "f1":
                    _testObject.SendKeys(Keys.F1);
                    break;
                case "f10":
                    _testObject.SendKeys(Keys.F10);
                    break;
                case "f11":
                    _testObject.SendKeys(Keys.F11);
                    break;
                case "f12":
                    _testObject.SendKeys(Keys.F12);
                    break;
                case "f2":
                    _testObject.SendKeys(Keys.F2);
                    break;
                case "f3":
                    _testObject.SendKeys(Keys.F3);
                    break;
                case "f4":
                    _testObject.SendKeys(Keys.F4);
                    break;
                case "f5":
                    _testObject.SendKeys(Keys.F5);
                    break;
                case "f6":
                    _testObject.SendKeys(Keys.F6);
                    break;
                case "f7":
                    _testObject.SendKeys(Keys.F7);
                    break;
                case "f8":
                    _testObject.SendKeys(Keys.F8);
                    break;
                case "f9":
                    _testObject.SendKeys(Keys.F9);
                    break;
                case "help":
                    _testObject.SendKeys(Keys.Help);
                    break;
                case "home":
                    _testObject.SendKeys(Keys.Home);
                    break;
                case "insert":
                    _testObject.SendKeys(Keys.Insert);
                    break;
                case "left":
                    _testObject.SendKeys(Keys.Left);
                    break;
                case "leftalt":
                    _testObject.SendKeys(Keys.LeftAlt);
                    break;
                case "leftcontrol":
                    _testObject.SendKeys(Keys.LeftControl);
                    break;
                case "leftshift":
                    _testObject.SendKeys(Keys.LeftShift);
                    break;
                case "meta":
                    _testObject.SendKeys(Keys.Meta);
                    break;
                case "multiply":
                    _testObject.SendKeys(Keys.Multiply);
                    break;
                case "null":
                    _testObject.SendKeys(Keys.Null);
                    break;
                case "numberpad0":
                    _testObject.SendKeys(Keys.NumberPad0);
                    break;
                case "numberpad1":
                    _testObject.SendKeys(Keys.NumberPad1);
                    break;
                case "numberpad2":
                    _testObject.SendKeys(Keys.NumberPad2);
                    break;
                case "numberpad3":
                    _testObject.SendKeys(Keys.NumberPad3);
                    break;
                case "numberpad4":
                    _testObject.SendKeys(Keys.NumberPad4);
                    break;
                case "numberpad5":
                    _testObject.SendKeys(Keys.NumberPad5);
                    break;
                case "numberpad6":
                    _testObject.SendKeys(Keys.NumberPad6);
                    break;
                case "numberpad7":
                    _testObject.SendKeys(Keys.NumberPad7);
                    break;
                case "numberpad8":
                    _testObject.SendKeys(Keys.NumberPad8);
                    break;
                case "numberpad9":
                    _testObject.SendKeys(Keys.NumberPad9);
                    break;
                case "pagedown":
                    _testObject.SendKeys(Keys.PageDown);
                    break;
                case "pageup":
                    _testObject.SendKeys(Keys.PageUp);
                    break;
                case "pause":
                    _testObject.SendKeys(Keys.Pause);
                    break;
                case "return":
                    _testObject.SendKeys(Keys.Return);
                    break;
                case "right":
                    _testObject.SendKeys(Keys.Right);
                    break;
                case "semicolon":
                    _testObject.SendKeys(Keys.Semicolon);
                    break;
                case "separator":
                    _testObject.SendKeys(Keys.Separator);
                    break;
                case "shift":
                    _testObject.SendKeys(Keys.Shift);
                    break;
                case "space":
                    _testObject.SendKeys(Keys.Space);
                    break;
                case "subtract":
                    _testObject.SendKeys(Keys.Subtract);
                    break;
                case "tab":
                    _testObject.SendKeys(Keys.Tab);
                    break;
                case "up":
                    _testObject.SendKeys(Keys.Up);
                    break;
                case "ctrl+c":
                    _testObject.SendKeys(Keys.Control + "c");
                    break;
                case "ctrl+v":
                    _testObject.SendKeys(Keys.Control + "v");
                    break;
                case "ctrl+a":
                    _testObject.SendKeys(Keys.Control + "a");
                    break;

            }

        }
        /// <summary>
        /// This will check all specified child checkboxes.
        /// </summary>
        /// <param name="dataContents">String[] : Labels of the checkboxes that have to be checked.</param>
        public void CheckMultiple(string[] dataContents)
        {
            //Get the Test Object.
            _testObject = WaitAndGetElement();

            //Get All Specified CheckBoxes
            List<IWebElement> allCheckBoxes = GetAllCheckBoxes(dataContents);

            //Checking All Specified Checkboxes.
            foreach (IWebElement checkbox in allCheckBoxes.Where(checkbox => !checkbox.Selected))
            {
                checkbox.Click();
            }
        }
        /// <summary>
        /// Get All checkbox Objects
        /// </summary>
        /// <param name="dataContents"></param>
        /// <returns></returns>
        private List<IWebElement> GetAllCheckBoxes(string[] dataContents)
        {
            //Get All Labels in the test object.
            IList<IWebElement> allLabels = _testObject.FindElements(By.TagName("label"));

            //Declaring List Of all Checkboxes that need to be processed.
            List<IWebElement> allCheckBoxes = new List<IWebElement>();

            //Processing each Labels one by one.
            foreach (IWebElement labels in allLabels)
            {
                //Checking if Checkbox need to be processed based upon inputs given.
                if (!checkLabels(dataContents, labels.Text))
                {
                    continue; //Continuing to next label.
                }

                IWebElement checkBoxOfLabel;
                //Fetching 'for' attribute of current label.
                var forAttribute = labels.GetAttribute("for");
                //There are checkbox groups in which for is not specified, these groups have different hierarchies.
                if (forAttribute != null)
                {
                    //Getting checkbox Element.
                    checkBoxOfLabel = Driver.FindElement(By.Id(forAttribute));

                    //Add checkbox to list.
                    allCheckBoxes.Add(checkBoxOfLabel);
                }
                else //If for attribute is not present in DOM for labels.
                {
                    //Getting CheckBox Element.
                    checkBoxOfLabel = labels.FindElement(By.TagName("input"));

                    //Add checkbox to list.
                    allCheckBoxes.Add(checkBoxOfLabel);
                }
            }
            return allCheckBoxes;
        }

        /// <summary>
        /// Uncheck all child checkboxes.
        /// </summary>
        /// <param name="dataContent">string[] : specified the labels of checkboxes.</param>
        public void UncheckMultiple(string[] dataContent)
        {
            //Get Test Object.
            _testObject = WaitAndGetElement();

            //Get all specified checkboxes.
            List<IWebElement> allCheckBoxes = GetAllCheckBoxes(dataContent);

            //UnChecking All Specified Checkboxes.
            foreach (IWebElement checkbox in allCheckBoxes)
            {
                if (checkbox.Selected)
                {
                    checkbox.Click();
                }
            }
        }

        /// <summary>
        /// Check if specified label is the current label.
        /// </summary>
        /// <param name="dataContent"></param>
        /// <param name="labelText"></param>
        /// <returns></returns>
        private bool checkLabels(string[] dataContent, string labelText)
        {
            return dataContent.Any(data => data.Trim().Equals(labelText));
        }

        /// <summary>
        ///  Method to check on checkbox and radio button.
        /// </summary>
        public void Check(string checkStatus = "")
        {
            try
            {
                _testObject = WaitAndGetElement();
                bool isSelected = _testObject.Selected;

                switch (checkStatus.ToLower().Trim())
                {
                    case "on":
                        if (!isSelected)
                        {
                            _testObject.Click();
                        }
                        break;
                    case "off":
                        if (isSelected)
                        {
                            _testObject.Click();
                        }
                        break;
                    default:
                        if (!isSelected)
                        {
                            _testObject.Click();
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
                        ExecuteScript(_testObject, "arguments[0].click();");
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
            Check("off");
        }


        /// <summary>
        /// Method to type specied text in text box.
        /// </summary>
        /// <param name="text">string : Text to type on text box object</param>
        public void SendKeys(string text)
        {
            try
            {

                _testObject = WaitAndGetElement();

                //Check ON or OFF condition for radio button or checkbox.
                if (text.Equals("ON", StringComparison.CurrentCultureIgnoreCase) || (text.Equals("OFF", StringComparison.CurrentCultureIgnoreCase)))
                {
                    var objType = _testObject.GetAttribute("type");
                    if (objType.Equals("checkbox", StringComparison.CurrentCultureIgnoreCase) ||
                         objType.Equals("radio", StringComparison.CurrentCultureIgnoreCase))
                    {
                        switch (text.ToLower())
                        {
                            case "on":
                                Property.StepDescription = "Check " + _testObject.Text;
                                Check();
                                break;
                            case "off":
                                Property.StepDescription = "Uncheck " + _testObject.Text;
                                UnCheck();
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
                        Driver.SwitchTo().Frame(_testObject); //Selecting the frame
                        Driver.FindElement(By.CssSelector("body")).SendKeys("text"); //entering the data to frame.
                    }
                    else
                    {
                        _testObject.Clear();
                        try
                        {
                            _testObject.SendKeys(text);
                        }
                        catch (ElementNotVisibleException enve)
                        {
                            try
                            {
                                ExecuteScript(_testObject, string.Format("arguments[0].value='{0}';", text));
                            }
                            catch
                            {
                                throw enve;
                            }
                        }
                        catch (Exception)
                        {
                            if (_objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                                throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0069").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute)); // added by 
                            throw;
                        }
                    }
                }
            }

            catch (StaleElementReferenceException)
            {
                _testObject = WaitAndGetElement();
                _testObject.Clear();
                try
                {
                    _testObject.SendKeys(text);
                }
                catch (Exception)
                {
                    if (_objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                        throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0069").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute)); // added by 
                    throw;
                }
            }
            catch (Exception e)
            {
                if (e.Message.Contains("Element is no longer attached to the DOM"))
                {
                    _testObject = WaitAndGetElement();
                    _testObject.Clear();
                    try
                    {
                        _testObject.SendKeys(text);
                    }
                    catch (Exception)
                    {
                        if (_objDataRow.ContainsKey(KryptonConstants.TEST_OBJECT))
                            throw new NoSuchElementException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0069").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute)); // added by 
                        throw;
                    }
                }
                else
                {
                    throw;
                }
            }
        }


        /// <summary>
        ///Method to submit form data.
        /// </summary>
        public void Submit()
        {
            _testObject = WaitAndGetElement();
            _testObject.Submit();
        }


        /// <summary>
        ///Method to fire specified event.
        /// </summary>
        public void FireEvent(string eventName)
        {
            //Retrieve test object from application
            _testObject = WaitAndGetElement();

            //Get all events that needs to be fired
            string[] arrEvents = eventName.Split('|');
            string[] arrScripts = new string[arrEvents.Length];


            //Start a for loop for each event, created javascript and store to script array
            for (int i = 0; i < arrEvents.Length; i++)
            {
                //Retrieve event name from array one by one
                eventName = arrEvents[i].Trim().ToLower(); //event must be in lower case.

                // Replace special characters
                eventName = Utility.ReplaceSpecialCharactersInString(eventName);

                //Eventname should start with "on" for internet explorer
                string onEventName;
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
                var script = "var canBubble = false;" + Environment.NewLine +
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
                if (!Browser.BrowserName.Equals("ie"))
                {
                    script = "var evt = document.createEvent(\"HTMLEvents\"); evt.initEvent(\"" +
                             eventName + "\", true, true );return !arguments[0].dispatchEvent(evt);";
                }

                //Store scripts in script specific array
                arrScripts[i] = script;

            }

            //Execute scripts now. This loop needs to be saparate as events needs to be fired as fast as possible
            foreach (string script in arrScripts)
            {
                ExecuteScript(_testObject, script);
            }
        }

        ///  <summary>
        /// Method to wait for specified object to become visible.
        ///  </summary>
        /// <param name="globalWaitTime">global wait time from the parameter file</param>
        /// <param name="optionDic">Dictionary : Option dictionary</param>
        /// <param name="waitTime">wait time for the object mention in the test case</param>
        /// <param name="modifier"></param>
        public string WaitForObject(string waitTime, string globalWaitTime, Dictionary<int, string> optionDic, string modifier)
        {
            try
            {
                int intWaitTime = 0;
                _modifier = modifier;
                _optiondic = optionDic;
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
                WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(intWaitTime));
                Func<IWebDriver, bool> condition = VerifyObjectDisplayCondition;
                DateTime dtBefore = DateTime.Now;
                wait.Until(condition);
                DateTime dtafter = DateTime.Now;
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
                throw new Exception("Object is not present on current page");
            }

        }

        ///  <summary>
        /// Method to wait for specified object to disappear.
        ///  </summary>
        /// <param name="optionDic">Dictionary : Option dictionary</param>
        public void WaitForObjectNotPresent(string waitTime, string globalWaitTime, Dictionary<int, string> optionDic)
        {
            if (waitTime.Equals(string.Empty))
            {
                waitTime = globalWaitTime;
            }
            _optiondic = optionDic;

            int intWaitTime = Int32.Parse(waitTime);
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(intWaitTime));
            Func<IWebDriver, bool> condition = VerifyObjectNotDisplayedCondition;
            wait.Until(condition);

        }

        ///  <summary>
        /// Wait for specified test object property to appear.
        ///  </summary>
        /// <param name="globalWaitTime">string : Global time out</param>
        ///  <param name="optionDic">Dictionary : Option Dictionary</param>
        public void WaitForObjectProperty(string propertyParam, string propertyValueParam, string globalWaitTime, Dictionary<int, string> optionDic)
        {
            _property = propertyParam.Trim();
            _propertyValue = propertyValueParam.Trim();
            _optiondic = optionDic;
            int intWaitTime = Int32.Parse(globalWaitTime);
            WebDriverWait wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(intWaitTime));
            Func<IWebDriver, bool> condition = VerifyObjectPropertyCondition;
            wait.Until(condition);
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
                return VerifyObjectProperty(_property, _propertyValue, _optiondic);
            }
            catch (NoSuchElementException)
            {
                return false;
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
                bool status = VerifyObjectDisplayed();
                if (!status && _modifier.Equals("refresh"))
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
                return !VerifyObjectDisplayed();
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
            _selenium = Browser.GetSeleniumOne();
            _selenium.DoubleClick(Attribute);
        }


        /// <summary>
        /// Method to select specified item from list.
        /// </summary>
        public void SelectItem(string[] itemList, bool selectMultiple = false)
        {
            string availbleItemsInList = string.Empty;
            string expectedItems = string.Empty;

            try
            {
                string firstItem = itemList.First();
                expectedItems = itemList.Aggregate(expectedItems, (current, item) => current + string.Format("\t" + item));
                _testObject = WaitAndGetElement();
                if (_testObject != null)
                    availbleItemsInList = _testObject.Text;


                #region  KRYPTON0156: Handling conditions where selectItem can be used for radio groups
                if (_testObject != null && (_testObject.GetAttribute("type") != null && _testObject.GetAttribute("type").ToLower().Equals("radio")))
                {
                    string radioValue = string.Empty;

                    //Check if given option is also stored in mapping column
                    string valueFromMapping = string.Empty;

                    foreach (KeyValuePair<string, string> mappingKey in _dicMapping)
                    {
                        if (mappingKey.Key.ToLower().Equals(firstItem.ToLower()))
                        {
                            valueFromMapping = mappingKey.Value;
                            break;
                        }
                    }

                    foreach (IWebElement radioObject in _testObjects)
                    {
                        //Retrieve value property of radio button and compare with passed text
                        string tempRadioValue = radioObject.GetAttribute("value");

                        if (tempRadioValue.ToLower().Equals(firstItem.ToLower()) || tempRadioValue.ToLower().Equals(valueFromMapping.ToLower()))
                        {
                            radioValue = tempRadioValue;
                            _testObject = radioObject;
                            break;
                        }
                    }

                    //Determine if a radio option was found or not, and take action based on that
                    if (!radioValue.Equals(string.Empty))
                    {
                        ExecuteScript(_testObject, "arguments[0].click();");
                        return;
                    }
                    Property.Remarks = string.Empty;
                    throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0022").Replace("{MSG}", firstItem));
                }
                #endregion

                // Entire section commented above works just fine, but damn slow on IE
                //Following code section is a shortcut for the same, and works reliably and fast on all browsers

                //Clear previus selection if multiple options needs to be selected

                if (selectMultiple)
                {
                    SelectElement select = new SelectElement(_testObject);
                    select.DeselectAll();
                }
                bool isSelectionSuccessful = false;
                foreach (string optionForSelection in itemList)
                {
                    string optionText = optionForSelection.Trim();


                    if (_testObject != null)
                    {
                        IList<IWebElement> allOptions = _testObject.FindElements(By.TagName("option"));
                        foreach (IWebElement option in allOptions)
                        {
                            try
                            {
                                string propertyValue = option.GetAttribute("value");
                                if (option.Displayed)
                                {
                                    if (option.Text.Contains(optionText))
                                    {
                                        isSelectionSuccessful = true;
                                        option.Click();
                                        break;
                                    }
                                    if (propertyValue != null && propertyValue.ToLower().Contains(optionText))
                                    {
                                        isSelectionSuccessful = true;
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
                                    isSelectionSuccessful = true;
                                    option.Click();
                                    break;
                                }
                                throw new Exception(string.Format("Could not select \"{0}\" in the list with options:\" {1} \". Actual error:{2}", optionText, availbleItemsInList, ex.Message));
                            }
                        }
                    }
                }
                if (!isSelectionSuccessful)
                {
                    throw new Exception(string.Format("Could not found \"{0}\" in the list with options:\" {1} \".", expectedItems, availbleItemsInList));

                }
            }
            catch
            {
                throw new Exception(string.Format("Could not select \"{0}\" in the list with options:\" {1} \".", expectedItems, availbleItemsInList));

            }

        }



        /// <summary>
        /// Method to select option from list by index.
        /// </summary>
        /// <param name="index">string : index of item to be select</param>
        public string SelectItemByIndex(string index)
        {
            int intIndex = Int32.Parse(index) - 1;
            _testObject = WaitAndGetElement();

            #region  KRYPTON0156: Handling conditions where selectItemByIndex can be used for radio groups
            if (_testObject.GetAttribute("type").ToLower().Equals("radio"))
            {
                _testObject = _testObjects.ElementAt(intIndex);
                ExecuteScript(_testObject, "arguments[0].click();");
                return _testObject.GetAttribute("value");
            }
            #endregion

            SelectElement select = new SelectElement(_testObject);

            @select.SelectByIndex(intIndex);
            IWebElement[] options = @select.Options.ToArray();
            return options[intIndex].Text;
        }


        /// <summary>
        /// Method to return test object property with specified property type.
        /// </summary>
        /// <param name="property">string : Name of the property</param>
        /// <returns>string : Property value</returns>
        public string GetObjectProperty(string property)
        {
            _testObject = WaitAndGetElement();
            string actualPropertyValue;
            string javascript;

            switch (property.ToLower())
            {
                case "text":
                    actualPropertyValue = _testObject.Text;
                    break;
                // Edited to work with firefox
                case "style.background":
                case "style.backgroundimage":
                case "style.background-image":
                    switch (Browser.BrowserName.ToLower())
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
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;

                case "style.color":
                    switch (Browser.BrowserName.ToLower())
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
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "style.fontweight":
                case "style.font-weight":
                    switch (Browser.BrowserName.ToLower())
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
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "style.left":
                    switch (Browser.BrowserName.ToLower())
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
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "disabled":
                    javascript = "return arguments[0].disabled;";
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;

                case "tooltip":
                case "title":
                    javascript = "return title=arguments[0].title;";
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "tag":
                case "html tag":
                case "htmltag":
                case "tagname":
                    actualPropertyValue = _testObject.GetAttribute("tagName");
                    break;
                case "scrollheight":
                    javascript = "return arguments[0].scrollHeight;";
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "filename":
                    javascript = "var src = arguments[0].src; src=src.split(\"/\"); return(src[src.length-1]);";
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "height":
                    javascript = "return arguments[0].height;";
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "width":
                    javascript = "return arguments[0].width;";
                    actualPropertyValue = ExecuteScript(_testObject, javascript);
                    break;
                case "checked":
                    javascript = "return arguments[0].checked;";
                    actualPropertyValue = ExecuteScript(_testObject, javascript).ToLower();
                    break;
                case "orientation":
                    // For orientation, calculate width and height
                    int height = Convert.ToInt16(GetObjectProperty("height"));
                    int width = Convert.ToInt16(GetObjectProperty("width"));

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
                    if (_testObject.GetAttribute("type").ToLower().Equals("radio"))
                    {
                        //Out of radio group, retrieve which one is selected
                        string radioValue = string.Empty;
                        javascript = "return arguments[0].checked;";

                        foreach (IWebElement radioObject in _testObjects)
                        {
                            radioValue = ExecuteScript(radioObject, javascript);

                            if (radioValue.ToLower().Equals("true") || radioValue.ToLower().Equals("on") || radioValue.ToLower().Equals("1"))
                            {
                                radioValue = radioObject.GetAttribute("value");
                                _testObject = radioObject;
                                break;
                            }
                            radioValue = "{no option selected}";
                        }

                        //Check if given option is also stored in mapping column, if so assign to actual property value
                        actualPropertyValue = radioValue;

                        foreach (KeyValuePair<string, string> mappingKey in _dicMapping)
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
                        actualPropertyValue = ExecuteScript(_testObject, javascript).ToLower();
                    }
                    break;
                default:
                    actualPropertyValue = _testObject.GetAttribute(property);
                    if (actualPropertyValue == null)
                    {
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0023").Replace("{MSG1}", property).Replace("{MSG2}", _testObject.Text));
                    }
                    break;
            }
            return actualPropertyValue;
        }

        /// <summary>
        ///Method to verify Object visiblity.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyObjectDisplayed()
        {
            try
            {
                IWebElement testObject = (IWebElement)GetElement(Driver);

                bool status = true;
                if (testObject == null || !testObject.Displayed)
                {
                    Property.Remarks = "Object is not displaying.";
                    status = false;
                }
                return status;
            }
            catch (NoSuchElementException)
            {
                throw;
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf(Exceptions.ERROR_ELEMENT_NOT_ATTACHED, StringComparison.OrdinalIgnoreCase) >= 0) // done Temp: for Chromedriver issue.
                {
                    throw new Exception("Object is not displaying.");
                }
                if (e.Message.IndexOf("unexpected alert open", StringComparison.OrdinalIgnoreCase) >= 0) // done Temp: for Chromedriver V2.2 issue.
                {
                    throw new Exception("Object is not displaying due to unexpected alert open.");
                }
                throw;
            }
        }


        /// <summary>
        /// Method toverify list item is present.
        /// </summary>
        /// <param name="listItemName">string : option to verify</param>
        /// <returns>boolean value</returns>
        public bool VerifyListItemPresent(string listItemName)
        {
            _testObject = WaitAndGetElement();
            SelectElement select = new SelectElement(_testObject);
            IList<IWebElement> option = @select.Options;
            bool flag = option.Any(element => element.Text.Equals(listItemName));
            if (!flag)
            {
                Property.Remarks = "List Item \"" + listItemName + "\" does not match with \"" + _testObject.Text + "\" values in list";
            }
            return flag;
        }




        /// <summary>
        /// Method to verify list itme is not present in list.
        /// </summary>
        /// <param name="listItemName">string : option to verify</param>
        /// <returns>boolean value</returns>
        public bool VerifyListItemNotPresent(string listItemName)
        {
            bool status = VerifyListItemPresent(listItemName);
            Property.Remarks = string.Empty;
            if (status)
            {
                Property.Remarks = "List Item \"" + listItemName + "\" is present";
            }
            return !status;
        }

        /// <summary>
        /// Method to verify object not present in web page.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyObjectPresent()
        {
            try
            {
                _testObject = WaitAndGetElement();
                if (_testObject != null)
                {
                    IWebElement element = (IWebElement)_testObject;
                    bool status = element.Displayed;
                    if (!status)
                    {
                        Property.Remarks = "Element is present but not displayed.";
                    }
                    return status;

                }
                Property.Remarks = "Element is not present.";
                return false;
            }
            catch (NoSuchElementException e)
            {
                Property.Remarks = e.Message;
                return false;
            }
        }



        /// <summary>
        ///Method to verify test object is present.
        /// </summary>
        /// <returns>boolean value</returns>
        public bool VerifyObjectNotPresent()
        {
            return WaitNonPresenceAndGetElement();
        }


        ///  <summary>
        /// Method to verify object property with specified property.
        ///  </summary>
        ///  <param name="property">string : Property name</param>
        /// <param name="propertyValue"></param>
        /// <param name="keywordDic"></param>
        /// <returns>boolean value</returns>
        public bool VerifyObjectProperty(string property, string propertyValue, Dictionary<int, string> keywordDic)
        {
            string actualPropertyValue = GetObjectProperty(property);
            var isKeyVerified = !keywordDic.Count.Equals(0) ? Utility.DoKeywordMatch(propertyValue, actualPropertyValue) : actualPropertyValue.Equals(propertyValue);
            if (!isKeyVerified)
            {
                Property.Remarks = "Property -\"" + property + "\", actual value \"" + actualPropertyValue + "\" does not match with expected value - \"" + propertyValue + "\".";
            }
            else
                Property.Remarks = "Property -\"" + property + "\", actual value \"" + actualPropertyValue + "\" match with expected value - \"" + propertyValue + "\".";

            return isKeyVerified;
        }

        public bool VerifyObjectPropertyNot(string property, string propertyValue, Dictionary<int, string> KeywordDic)
        {
            bool status = VerifyObjectProperty(property, propertyValue, KeywordDic);
            Property.Remarks = string.Empty;
            if (status)
            {
                Property.Remarks = "Property \"" + property + "\" matches with actual value \"" + propertyValue;
            }
            return !status;
        }

        /// <summary>
        ///Method to hightlight current test object in web page.
        /// </summary>
        private void HightlightObject(IWebElement tObject = null)
        {

            if (DebugMode.Equals("true", StringComparison.OrdinalIgnoreCase)
                && Browser.IsRemoteExecution.ToString().Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                string propertyName = "outline";
                string originalColor = "none";
                string highlightColor = "#00ff00 solid 3px";

                try
                {
                    //This works with internet explorer
                    originalColor = ExecuteScript(tObject, "return arguments[0].currentStyle." + propertyName);
                }
                catch
                {
                    //This works with firefox, chrome and possibly others
                    originalColor = ExecuteScript(tObject, "return arguments[0].style." + propertyName);
                }

                ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + highlightColor + "'");
                Thread.Sleep(50);
                ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + originalColor + "'");
                Thread.Sleep(50);
                ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + highlightColor + "'");
                Thread.Sleep(50);
                ExecuteScript(tObject, "arguments[0].style." + propertyName + " = '" + originalColor + "'");
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
                object scriptRet;
                if (Browser.IsRemoteExecution.Equals("false"))
                {
                    switch (Browser.BrowserName.ToLower())
                    {
                        case KryptonConstants.BROWSER_FIREFOX:
                            FirefoxDriver ffscriptdriver = (FirefoxDriver)Driver;
                            scriptRet = ffscriptdriver.ExecuteScript(scriptToExecute, tObject);
                            if (scriptRet != null)
                            {
                                return scriptRet.ToString();
                            }
                            break;
                        case KryptonConstants.BROWSER_IE:
                            InternetExplorerDriver iescriptdriver = (InternetExplorerDriver)Driver;

                            scriptRet = iescriptdriver.ExecuteScript(scriptToExecute, tObject);
                            if (scriptRet != null)
                            {
                                return scriptRet.ToString();
                            }
                            break;
                        case KryptonConstants.BROWSER_CHROME:
                            ChromeDriver crscriptdriver = (ChromeDriver)Driver;
                            scriptRet = crscriptdriver.ExecuteScript(scriptToExecute, tObject);
                            if (scriptRet != null)
                            {
                                return scriptRet.ToString();
                            }
                            break;
                        case KryptonConstants.BROWSER_SAFARI:
                            SafariDriver safacriptdriver = (SafariDriver)Driver;
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
                    RemoteWebDriver rdscriptriver = (RemoteWebDriver)Driver;
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
                if (_testObject != null && _objDataRow.Count > 0)
                    throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0068").Replace("{MSG3}", _objDataRow[KryptonConstants.TEST_OBJECT]).Replace("{MSG4}", _objDataRow["parent"]).Replace("{MSG1}", AttributeType).Replace("{MSG2}", Attribute));
                throw new Exception("Java Script Error :  " + e.Message);
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
                if (AttributeType == null || AttributeType.Equals(string.Empty))
                {
                    _testObject = null;
                }
                else
                {
                    _testObject = GetElement(Driver);

                }
                return ExecuteScript(_testObject, scriptToExecute);
            }
            catch (Exception e)
            {
                Browser.SwitchToMostRecentBrowser();
                return ExecuteScript(_testObject, scriptToExecute);
            }
        }

        /// <summary>
        ///Method to set property present in DOM
        /// </summary>
        /// <param name="property">string : Name of property</param>
        /// <param name="propertyValue">string : Value of property</param>
        public void SetAttribute(string property, string propertyValue)
        {
            _testObject = WaitAndGetElement();
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
            ExecuteScript(_testObject, javascript);
        }

        /// <summary>
        /// Method to drag the source object on the target object.
        /// </summary>
        /// <param name="targetObjDic">Dictionary : Target object details</param>
        public void DragAndDrop(Dictionary<string, string> targetObjDic)
        {
            IWebElement sourceObject = WaitAndGetElement();

            // Set second object's information in dictionary
            SetObjDataRow(targetObjDic);
            IWebElement targetObject = WaitAndGetElement();

            Actions action = new Actions(Driver);
            action.DragAndDrop(sourceObject, targetObject).Perform();
            Thread.Sleep(100);
        }

        /// <summary>
        /// Upload the file
        /// </summary>
        /// <param name="path"></param>
        /// <param name="element"></param>
        public bool UploadFile(string path, string element)
        {
            string p = @"" + path;
            if (!File.Exists(p))
            {
                return false;
            }
            InternetExplorerDriver iedriver = (InternetExplorerDriver)Driver;
            iedriver.FindElement(By.Id(element)).SendKeys(p);
            return true;
        }
    }
}
