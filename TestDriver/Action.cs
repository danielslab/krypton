/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Krypton.TestEngine.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Action mapping class
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Common;
using System.Reflection;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using AutoItX3Lib;
using System.Text.RegularExpressions;
using Driver.Browsers;

namespace TestDriver
{

    public class Action
    {
        [DllImport("AutoItX3.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static public extern int AU3_MouseUp([MarshalAs(UnmanagedType.LPStr)] string button);
        private string _stepAction = string.Empty;
        private string _parent = string.Empty;
        private string _testObject = string.Empty;
        private string _testData = string.Empty;
        public static string BrowserVersion = string.Empty;
        private readonly Dictionary<int, string> _keywordDic = new Dictionary<int, string>();
        public static Dictionary<string, string> ObjDataRow = new Dictionary<string, string>();
        public static Dictionary<string, string> ObjSecondDataRow = new Dictionary<string, string>();
        private Browser _objBrowser = null;
        private Driver.ITestObject _objTestObject = null;
        private Driver.RecoveryScenarios _objRecovery = null;
        private DialogHandler _objHandler = null;
        private IProjectMethods _pluginProject = null;
        private string _verificationMessage = string.Empty;
        private string _snapShotOption = Property.SnapshotOption.ToLower();
        private readonly string _globalTimeout = Property.GlobalTimeOut;
        private string _debugMode = Property.DebugMode;
        private readonly DataSet _datasetRecoverBrowser;
        private readonly DataSet _datasetRecoverPopup;
        private readonly DataSet _datasetOr;

        // Declared a Win32 function to control other processes' window.
        [DllImport("user32.dll")]
        static extern bool ShowWindow(int hWnd, int nCmdShow);

        /// <summary>
        ///Default constructor initialize TestObject and Browser class
        /// </summary>
        public Action(DataSet recoverPopupData, DataSet recoverBrowserData, DataSet orData)
        {
            _datasetRecoverPopup = recoverPopupData;
            _datasetRecoverBrowser = recoverBrowserData;
            _datasetOr = orData;
            string browserName = Utility.GetParameter(Property.BrowserString).ToLower();

            _objBrowser = new Browser(Property.ErrorCaptureAs);
            _objTestObject = new Driver.TestObject(Utility.GetParameter("ObjectTimeout"));

            string[] availablePlugins = Directory.GetFiles(Property.ApplicationPath, "*ProjectPlugin.dll");
            if (availablePlugins.Length > 0)
            {
                foreach (string availablePlugin in availablePlugins)
                {
                    Assembly asm = Assembly.LoadFrom(availablePlugin);
                    Type myType = asm.GetType(asm.GetName().Name + ".MatchProjectPlugin");
                    _pluginProject = (IProjectMethods)Activator.CreateInstance(myType);

                }

            }
            //Initialize Object for specified language to create Selenium script.        
            switch (Property.ScriptLanguage.ToLower())
            {
                case "java":
                    break;
                case "php":
                    break;
            }

        }

        /// <summary>
        /// Saving the script 
        /// </summary>
        public static void SaveScript()
        {
        }

        #region Call to action in driver according to action step

        ///  <summary>
        /// This method will call action from Driver Class. 
        ///  </summary>
        ///  <param name="action">Step action to perfrom by driver</param>
        ///  <param name="parent=">Parent Object</param>
        /// <param name="parent"></param>
        /// <param name="child">Test object on which operation to be perform</param>
        ///  <param name="data">Test Data</param>
        ///  <param name="modifier"></param>
        ///  <returns></returns>
        public void Do(string action, string parent = null, string child = null, string data = null, string modifier = "")
        {
            _snapShotOption = Property.SnapshotOption.ToLower();
            //check for locator directly in test case.
            try
            {
                if (child != null && child.Contains("="))
                {
                    string how = child.Split('=')[0].ToLower().Trim();
                    string what = child.Replace(child.Split('=')[0] + "=", string.Empty);
                    ObjDataRow[KryptonConstants.HOW] = how;
                    ObjDataRow[KryptonConstants.WHAT] = what.Trim();
                    ObjDataRow[KryptonConstants.LOGICAL_NAME] = string.Empty;
                    ObjDataRow[KryptonConstants.OBJ_TYPE] = string.Empty;
                    ObjDataRow[KryptonConstants.MAPPING] = string.Empty;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            _stepAction = action;
            _testObject = child;
            _testData = data;
            //parse modifier
            modifier = modifier.ToLower().Trim();
            int i = 1;
            _keywordDic.Clear();//clearing previous keyword.

            string browserDimension = null; //browser dimensions
            for (int v = 0; ; v++)
            {
                if (modifier.Contains("{"))
                {
                    var stindex = modifier.IndexOf("{", StringComparison.Ordinal);
                    modifier = modifier.Remove(stindex, 1);
                    var endindex = modifier.IndexOf("}", StringComparison.Ordinal);
                    string keyVariable = modifier.Substring(stindex, (endindex - stindex));
                    if (keyVariable.ToLower().Contains("windowsize") || keyVariable.ToLower().Contains("window"))
                        browserDimension = keyVariable;
                    _keywordDic.Add(i, keyVariable);
                    i++;
                    stindex = modifier.IndexOf("}", StringComparison.Ordinal);
                    modifier = modifier.Remove(stindex, 1);
                }
                else
                {
                    break;
                }
            }

            if (_keywordDic.ContainsValue("nowait"))
                Property.NoWait = true;
            else
                Property.NoWait = false;

            string stepStatus = string.Empty;
            bool verification = true;

            Utility.DriverKeydic = null;
            Utility.DriverKeydic = _keywordDic;

            // Initializing Step Description to empty, this will allow individual cases to write description
            if (!modifier.Contains("recovery"))   // Consume last Step Description during recovery.
                Property.StepDescription = string.Empty;
            #region  Simplifying the parameters before passed to Actual actions.
            string[] dataContent = null;
            string contentFirst = string.Empty;
            string contentSecond = string.Empty;
            string contentThird = string.Empty;
            string contentFourth = string.Empty;
            string contentFifth = string.Empty;

            if (data != null)
            {
                dataContent = data.Split(Property.Seprator);
                switch (dataContent.Length)
                {
                    case 1:
                        contentFirst = dataContent[0];
                        break;
                    case 2:
                        contentFirst = dataContent[0];
                        contentSecond = dataContent[1];
                        break;
                    case 3:
                        contentFirst = dataContent[0];
                        contentSecond = dataContent[1];
                        contentThird = dataContent[2];
                        break;
                    case 4:
                        contentFirst = dataContent[0];
                        contentSecond = dataContent[1];
                        contentThird = dataContent[2];
                        contentFourth = dataContent[3];
                        break;
                    case 5:
                        contentFirst = dataContent[0];
                        contentSecond = dataContent[1];
                        contentThird = dataContent[2];
                        contentFourth = dataContent[3];
                        contentFifth = dataContent[4];
                        break;
                }
                contentFirst = Utility.ReplaceSpecialCharactersInString(contentFirst.Trim());
                contentSecond = Utility.ReplaceSpecialCharactersInString(contentSecond.Trim());
                contentThird = Utility.ReplaceSpecialCharactersInString(contentThird.Trim());
                contentFourth = Utility.ReplaceSpecialCharactersInString(contentFourth.Trim());
                contentFifth = Utility.ReplaceSpecialCharactersInString(contentFifth.Trim());
            }
            #endregion
            try
            {
                _objBrowser.SetObjDataRow(ObjDataRow);
                _objTestObject.SetObjDataRow(ObjDataRow, _stepAction);
                _objHandler = new DialogHandler();
                switch (_stepAction.ToLower())
                {
                    case "waitforpage":
                        _objTestObject.WaitForPage(contentFirst);
                        break;
                    //action are performed accoding to below action steps.
                    #region Action-> settestmode
                    case "settestmode":
                        // Assuming page object is given and Mode is paased as a data field. 
                        Property.NoWait = true;
                        string testModeVariable = Property.TestMode;
                        if (!string.IsNullOrWhiteSpace(contentSecond))
                        {
                            testModeVariable = Property.TestMode + "[" + contentFirst + "]";
                        }

                        //Update test mode only if object could be located
                        if (_objTestObject.VerifyObjectPresent())
                        {
                            Utility.SetVariable(testModeVariable, contentSecond.ToLower().Trim());
                        }

                        //Display test mode on console if debug mode is true
                        if (Utility.GetParameter("debugmode").ToLower().Equals("true"))
                            Property.Remarks = "Current Execution Test Mode =" + "'" + Utility.GetVariable(testModeVariable) + "'";
                        break;
                    #endregion
                    #region Action-> ShutDownDriver
                    case "shutdowndriver":
                        _objBrowser.Shutdown();
                        break;
                    #endregion
                    //Delete all internet temporery file.
                    #region Action-> ClearBrowserCache
                    case "clearbrowsercache":
                        Property.StepDescription = "Clear Browser Cache";
                        BrowserManager.Browser.ClearCache();
                        break;
                    #endregion
                    //Close most recent tab of browser associated with driver. 
                    #region Action-> CloseBrowser
                    case "closebrowser":
                        Property.StepDescription = "Close Browser";
                        _objBrowser.CloseBrowser();
                        break;
                    #endregion
                    //Close All Browsers
                    #region Action-> CloseAllBrowsers
                    case "closeallbrowsers":
                        if (Utility.GetParameter(Property.BrowserString).ToLower() != "android" && Utility.GetParameter(Property.BrowserString).ToLower() != "iphone" && Utility.GetParameter(Property.BrowserString).ToLower() != "selendroid")
                        {
                            Property.StepDescription = "Close all opened browsers.";
                            _objBrowser.CloseAllBrowser();
                        }
                        break;
                    #endregion
                    //Fire specified event on test object, eg.Click event.  
                    #region Action-> FireEvent
                    case "fireevent":
                        Property.NoWait = true;
                        Property.StepDescription = "Fire '" + contentFirst + "' event on " + _testObject;
                        _objTestObject.FireEvent(contentFirst);
                        break;
                    #endregion
                    //Start new instance of driver and open browser with specified url 
                    #region Action->OpenBrowser
                    case "openbrowser":
                        string browserName = Utility.GetParameter(Property.BrowserString).ToLower();
                        bool deleteCookie = !modifier.ToLower().Contains("keepcookies");
                        Property.StepDescription = "Open a new browser and navigate to url '" + contentFirst;
                        string remoteUrl = Property.RemoteUrl;
                        string isRemoteExecution = Property.IsRemoteExecution;
                        string profilePath = string.Empty;
                        // firefoxProfilePath parameter determines whether to use Firefox Profile or not.
                        switch (browserName)
                        {
                            case KryptonConstants.BROWSER_CHROME: profilePath = Utility.GetVariable("ChromeProfilePath");
                                break;
                            case KryptonConstants.BROWSER_FIREFOX: profilePath = Utility.GetVariable("FirefoxProfilePath");
                                break;
                            default: break;
                        }

                        //addonsPath parameter determines whether to load firefox addons or not.
                        string addonsPath = Utility.GetParameter("AddonsPath");

                        if (contentFirst.Equals(string.Empty))
                        {
                            contentFirst = Property.ApplicationUrl;
                        }

                        Exception openBrowserEx = null;
                        try
                        {

                            _objBrowser = Browser.OpenBrowser(browserName, deleteCookie, contentFirst,
                                                                        isRemoteExecution, remoteUrl, profilePath,
                                                                        addonsPath, _datasetRecoverPopup, _datasetRecoverBrowser, _datasetOr, browserDimension);
                        }
                        catch (Exception ex)
                        {
                            openBrowserEx = ex;
                            if (Property.IsRemoteExecution.ToLower().Equals("true"))
                                throw ex;

                        }

                        _objTestObject = new Driver.TestObject(Utility.GetParameter("ObjectTimeout"));
                        _objRecovery = new Driver.RecoveryScenarios(_datasetRecoverPopup, _datasetRecoverBrowser, _datasetOr, _objTestObject);

                        #region  Region containing JavaScript to maximize window and get Browser version string.
                        string browserVer;
                        Utility.SetVariable("BrowserVersion", BrowserVersion);
                        try
                        {
                            browserVer = _objTestObject.ExecuteStatement("return navigator.userAgent;");
                        }
                        catch (Exception)
                        {
                            throw openBrowserEx;
                        }
                        if (!browserVer.Equals(string.Empty))
                        {
                            switch (browserName.ToLower())
                            {
                                case KryptonConstants.BROWSER_FIREFOX:
                                    browserVer = browserVer.Substring(browserVer.IndexOf("Firefox/", StringComparison.Ordinal) + 8);
                                    if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                        browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                    BrowserVersion = browserVer;
                                    break;
                                case "ie":
                                case "iexplore":
                                case "internetexplorer":
                                case "internet explorer":
                                    BrowserVersion = browserVer.Substring(browserVer.IndexOf("MSIE ", StringComparison.Ordinal) + 5).Split(';')[0];
                                    string keyName = null;
                                    // Read the system registry to get the IE version in case tests are running locally
                                    if (!Utility.GetParameter("RunRemoteExecution").Equals("true", StringComparison.OrdinalIgnoreCase))
                                    {
                                        keyName = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer";
                                        BrowserVersion = (string)Registry.GetValue(keyName, "svcVersion", "key not found in Registry");
                                        if (BrowserVersion.Equals("key not found in Registry"))
                                        {
                                            BrowserVersion = (string)Registry.GetValue(keyName, "Version", "key not found in Registry");
                                            if (BrowserVersion.Equals("key not found in Registry")) // chek if version key is also not available in registry.
                                            {
                                                BrowserVersion = string.Empty;
                                            }
                                            else
                                                BrowserVersion = BrowserVersion.Split('.')[0] + "." + BrowserVersion.Split('.')[1];
                                        }
                                        else
                                        {
                                            BrowserVersion = BrowserVersion.Split('.')[0] + "." + BrowserVersion.Split('.')[1];
                                        }
                                    }
                                    break;
                                case KryptonConstants.BROWSER_CHROME:

                                    browserVer = browserVer.Substring(browserVer.IndexOf("Chrome/", StringComparison.Ordinal) + 7).Split(' ')[0];
                                    if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                        browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                    BrowserVersion = browserVer;
                                    break;
                                case KryptonConstants.BROWSER_OPERA:
                                    if (browserVer.Contains("Version/"))
                                        browserVer = browserVer.Substring(browserVer.IndexOf("Version/", StringComparison.Ordinal) + 8);
                                    else
                                        browserVer = browserVer.Substring(browserVer.IndexOf("Opera", StringComparison.Ordinal) + 6).Split(' ')[0];
                                    if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                        browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                    BrowserVersion = browserVer;
                                    break;
                                case KryptonConstants.BROWSER_SAFARI:
                                    browserVer = browserVer.Substring(browserVer.IndexOf("Version/", StringComparison.Ordinal) + 8).Split(' ')[0];
                                    if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                        browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                    BrowserVersion = browserVer;
                                    break;
                            }
                        }
                        #endregion

                        Utility.SetVariable("BrowserVersion", BrowserVersion);

                        _objBrowser.SetObjDataRow(ObjDataRow);
                        Utility.SetVariable(Property.BrowserVersion, BrowserVersion);
                        break;
                    #endregion
                    //Navigate to back on current web page
                    #region Action->GoBack
                    case "goback":
                        _objBrowser.GoBack();
                        Property.StepDescription = "Go backwards in the browser";
                        break;
                    #endregion
                    //Navigate to forward on current web page. :
                    #region Action->GoForward
                    case "goforward":
                        Property.StepDescription = "Go forward in the browser";
                        _objBrowser.GoForward();
                        break;
                    #endregion
                    //Refresh current web page. :
                    #region Action->RefreshBrowser
                    case "refreshbrowser":
                        Property.StepDescription = "Refresh browser";
                        _objBrowser.Refresh();
                        break;
                    #endregion
                    case "switchtorecentbrowser":
                    case "switchtonewbrowser":
                        Property.StepDescription = "Set focus to most recently opened window";
                        _objBrowser.SetBrowserFocus();
                        break;

                    //Navigate to new specified url in current web page.
                    #region Action->NavigateURL
                    case "navigateurl":
                        Property.StepDescription = "Navigate to url '" + contentFirst + "' in currently opened browser";
                        if (!contentFirst.StartsWith(@"http://"))
                            contentFirst = @"http://" + contentFirst;
                        _objBrowser.NavigationUrl(contentFirst);
                        break;
                    #endregion
                    //Clear existing data from test object. 
                    case "clear":
                    case "cleartext":
                        Property.StepDescription = "Clear object '" + _testObject + "'";
                        _objTestObject.ClearText();
                        break;
                    //Check radio button and check box.
                    case "check":
                        Property.StepDescription = "Check '" + contentFirst + "' checkbox '" + _testObject + "'";
                        _objTestObject.Check(contentFirst);
                        break;
                    //Uncheck radio button and check box. :
                    case "uncheck":
                        Property.StepDescription = "Uncheck checkbox '" + _testObject + "'";
                        _objTestObject.UnCheck();
                        break;
                    case "checkmultiple":
                        Property.StepDescription = "Check multiple checkboxes of '" + _testObject + "'";
                        _objTestObject.CheckMultiple(dataContent);
                        break;
                    case "uncheckmultiple":
                        Property.StepDescription = "Uncheck multiple checkboxes of '" + _testObject + "'";
                        string[] dataC = data.Split(Property.Seprator);
                        _objTestObject.UncheckMultiple(dataC);
                        break;
                    case "swipeobject":
                        Property.StepDescription = "swipe object in " + contentFirst + " direction " + _testObject;
                        Utility.SetVariable(_testObject, contentFirst);
                        break;
                    //Perform click action on associated test object. :
                    #region Action->Click
                    case "click":
                        Property.StepDescription = "Click on '" + _testObject + "'";

                        DateTime dtbefore = new DateTime();
                        if (ObjDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && ObjDataRow[KryptonConstants.OBJ_TYPE].ToLower().Contains("winbutton"))
                        {

                            dtbefore = DateTime.Now;
                            if (ObjDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && ObjDataRow[KryptonConstants.OBJ_TYPE].ToLower().Contains("winbutton") && Property.IsRemoteExecution.ToLower().Equals("true"))
                            {
                                // do nothing
                            }
                            else
                            {
                                var autoit = new AutoItX3();
                                string[] windowDetail = Regex.Split(ObjDataRow[KryptonConstants.WHAT], "//");

                                autoit.WinActivate(windowDetail[0], string.Empty);

                                autoit.ControlClick(windowDetail[0], string.Empty, windowDetail[1]);
                            }
                        }
                        else
                        {
                            _objTestObject.Click(_keywordDic, data);
                        }
                        break;

                    #endregion
                    case "doubleclick":
                        Property.StepDescription = "Double click on '" + _testObject + "'";
                        _objTestObject.DoubleClick();
                        break;
        #endregion
                    //Enter unique data in test object. 
                    #region Action-> EnterUniqueData
                    case "enteruniquedata":
                        Property.StepDescription = "Enter unique string of " + contentFirst + " characters in " + _testObject;
                        int length = Convert.ToInt16(contentFirst);
                        //passing length value if mentioned in test case.
                        string strUnique = Utility.GenerateUniqueString(length);
                        Utility.SetVariable(_testObject, strUnique);
                        _objTestObject.SendKeys(strUnique);
                        break;
                    #endregion
                    //Enter specified data in test object. :
                    case "enterdata":
                    case "type":

                    #region Action->TypeString
                    case "typestring":
                        if (contentFirst.Equals("ON", StringComparison.CurrentCultureIgnoreCase))
                            _stepAction = "Check";
                        if (contentFirst.Equals("OFF", StringComparison.CurrentCultureIgnoreCase))
                            _stepAction = "Unchek";

                        Property.StepDescription = "Enter text " + contentFirst + " in " + _testObject;

                        Utility.SetVariable(_testObject, contentFirst);//implicitely set key/Value to runtimedic dictionary before enter any data.
                        if (ObjDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && ObjDataRow[KryptonConstants.OBJ_TYPE].ToLower().Equals("winedit"))
                        {
                            if (ObjDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && ObjDataRow[KryptonConstants.OBJ_TYPE].ToLower().Contains("winedit") && Property.IsRemoteExecution.ToLower().Equals("true"))
                            {
                                UploadFileOnRemote oUploadFileOnRemote = new UploadFileOnRemote(Utility.GetParameter("browser"), Property.RemoteMachineIP, contentFirst.Trim());
                                oUploadFileOnRemote.UploadFileWithAutoIt();
                            }
                            else
                                _objHandler.EnterdataInDialog(contentFirst);
                        }
                        else
                        {
                            _objTestObject.SendKeys(contentFirst);
                        }
                        break;
                    #endregion
                    #region Action->UploadFileOnRemote
#if(DEBUG)  //used unly for dev testing purpose

                    case "uploadfileonremote":
                        if (Property.IsRemoteExecution.ToLower().Equals("true"))
                        {
                            Property.StepDescription = "Upload file " + contentFirst;
                            UploadFileOnRemote oUploadFileOnRemote = new UploadFileOnRemote(Utility.GetParameter("browser"), Property.RemoteMachineIP, contentFirst);
                            oUploadFileOnRemote.UploadFileWithAutoIt();
                        }
                        break;
#endif
                    #endregion
                    //Press non alphabetic key.
                    #region Action->KeyPress
                    case "keypress":
                        Property.StepDescription = "Press key " + contentFirst + " on " + _testObject;
                        _objTestObject.KeyPress(contentFirst);
                        break;
                    #endregion
                    //Adding submit in case this helps when clicks are missing.
                    #region Action->Submit
                    case "submit":
                        Property.StepDescription = "Submit " + _testObject;
                        _objTestObject.Submit();
                        break;
                    #endregion
                    //Perform wait action for specified time duration.
                    case "pause":
                    #region Action->Wait
                    case "wait":
                        Property.StepDescription = "Pause execution for " + contentFirst + " seconds";
                        browserName = Utility.GetParameter(Property.BrowserString).ToLower();
                        _objBrowser.Wait(contentFirst);
                        break;
                    #endregion
                    //Select option form List. 
                    #region Action->SelectItem
                    case "selectitem":
                        Property.StepDescription = "Select '" + contentFirst + "' from " + _testObject;
                        Utility.SetVariable(_testObject, contentFirst);
                        _objTestObject.SelectItem(dataContent);
                        break;
                    #endregion
                    //Select multiple options form List.
                    case "selectmultipleitems":
                    case "selectmultipleitem":
                    #region Action->SelectItems
                    case "selectitems":
                        Property.StepDescription = "Select multiple items '" + data + "' from " + _testObject;
                        Utility.SetVariable(_testObject, data);
                        _objTestObject.SelectItem(dataContent, true);
                        break;
                    #endregion
                    //Select option from List on the besis of index. 
                    #region Action-> SelectItemByIndex
                    case "selectitembyindex":
                        Property.StepDescription = "Select " + contentFirst + "th item from " + _testObject;
                        string optionValue = _objTestObject.SelectItemByIndex(contentFirst);
                        Utility.SetVariable(_testObject, optionValue);
                        break;
                    #endregion
                    //Wait for a specified condition to be happened on test object. 
                    #region Action->WaitForObject
                    case "waitforobject":
                        Property.StepDescription = "Wait until " + _testObject + " becomes available";
                        string actualwaitTime = _objTestObject.WaitForObject(contentFirst, _globalTimeout, _keywordDic, modifier.ToLower());
                        break;
                    #endregion
                    //Wait for a specified condition to be happened on test object. 
                    #region Action->WaitForObjectNotPresent
                    case "waitforobjectnotpresent":
                        Property.StepDescription = "Wait until " + _testObject + " disappears";
                        _objTestObject.WaitForObjectNotPresent(contentFirst, _globalTimeout, _keywordDic);
                        break;
                    #endregion
                    //Wait for specified property to enable. :
                    #region Action->WaitForProperty
                    case "waitforproperty":
                        Property.StepDescription = "Wait for " + _testObject + " to achieve value '" + contentSecond +
                                                    "' for '" + contentFirst + "' property";
                        _objTestObject.WaitForObjectProperty(contentFirst, contentSecond, _globalTimeout, _keywordDic);
                        break;
                    #endregion
                    //Accept alert :
                    #region Action->AcceptAlert
                    case "acceptalert":
                        Property.StepDescription = "Accept alert";
                        _objBrowser.AlertAccept();
                        break;
                    #endregion
                    #region Action->DismissAlert
                    case "dismissalert":
                        _objBrowser.AlertDismiss();
                        break;
                    #endregion
                    //Verify text on alert page. :
                    #region Action->VerifyAlertText
                    case "verifyalerttext":
                        Property.StepDescription = "Verify alert contains text '" + contentFirst + "'";
                        verification = _objBrowser.VerifyAlertText(contentFirst, _keywordDic);
                        break;
                    #endregion
                    //Get value of specified property of test object. :
                    case "getattribute":
                    #region Action->GetObjectProperty
                    case "getobjectproperty":
                        string property = contentFirst.Trim();
                        string propertyVariable = _testObject + "." + property;
                        if (!contentSecond.Equals(String.Empty))
                            propertyVariable = contentSecond.Trim();
                        string propertyValue = _objTestObject.GetObjectProperty(property);
                        Utility.SetVariable(propertyVariable, propertyValue);
                        break;
                    #endregion
                    //set property present in DOM
                    #region Action->SetAttribute
                    case "setattribute":

                        property = contentFirst.Trim();
                        propertyValue = contentSecond.Trim();
                        _objTestObject.SetAttribute(property, propertyValue);
                        break;
                    #endregion
                    //Get value of specified property of current web page. :

                    #region Action->GetPageProperty
                    case "getpageproperty":
                        propertyValue = _objBrowser.GetPageProperty(contentFirst.Trim());
                        propertyVariable = parent + "." + contentFirst.Trim();
                        if (!contentSecond.Equals(String.Empty))
                            propertyVariable = contentSecond.Trim();
                        Utility.SetVariable(propertyVariable, propertyValue);
                        break;
                    #endregion
                    //Verify that specified text is present on current web page or test object. :                    
                    case "verifytextcontained":
                    #region Action->VerifyTextPresent
                    case "verifytextpresent":
                        //- In case of text on page, Dictionary will be empty ie- count= 0.
                        if (ObjDataRow.Count == 0)
                        {
                            verification = _objBrowser.VerifyTextPresentOnPage(contentFirst, _keywordDic);
                        }
                        else
                        {
                            verification = _objTestObject.VerifyObjectProperty("text", contentFirst, _keywordDic);
                        }
                        break;
                    #endregion
                    //Verify that specified text is present on current web page or test object. :

                    case "verifytextnotcontained":
                    #region Action->VerifyTextNotPresent
                    case "verifytextnotpresent":

                        if (ObjDataRow[KryptonConstants.HOW].Equals(string.Empty) || ObjDataRow[KryptonConstants.WHAT].Equals(string.Empty) || ObjDataRow[KryptonConstants.HOW] == null || ObjDataRow[KryptonConstants.WHAT] == null || ObjDataRow[KryptonConstants.HOW].ToLower().Equals("url"))
                        {
                            verification = !_objBrowser.VerifyTextPresentOnPage(contentFirst, _keywordDic);
                        }
                        else
                        {
                            verification = _objTestObject.VerifyObjectPropertyNot("text", contentFirst, _keywordDic);
                        }
                        break;
                    #endregion
                    //Verify that specified text is present on current web page. :
                    #region Action->VerifyTextOnPage
                    case "verifytextonpage":
                        verification = _objBrowser.VerifyTextPresentOnPage(contentFirst, _keywordDic);
                        if (verification)
                            Property.StepDescription = "The Text : " + contentFirst + " was present on the page";
                        else
                            Property.StepDescription = "The Text : " + contentFirst + " was NOT present on the page";
                        break;
                    #endregion
                    //Verify that specified text is not present on current web page. :
                    #region Action-> VerifyTextNotOnPage
                    case "verifytextnotonpage":
                        verification = !_objBrowser.VerifyTextPresentOnPage(contentFirst, _keywordDic);
                        if (verification)
                        {
                            stepStatus = ExecutionStatus.Pass;
                            Property.Remarks = "Text  : \"" + contentFirst + "\" is not found on current Page.";
                        }
                        else
                        {
                            stepStatus = ExecutionStatus.Fail;
                            Property.Remarks = "Text  : \"" + contentFirst + "\" is found on current Page.";
                        }
                        break;
                    #endregion
                    //Verify that List option is persent in focused List. :
                    case "verifylistitempresent":
                        verification = _objTestObject.VerifyListItemPresent(contentFirst);
                        break;
                    //Verify that List option is not persent in focused List. :
                    case "verifylistitemnotpresent":
                        verification = _objTestObject.VerifyListItemNotPresent(contentFirst);
                        break;
                    //Verify that specified test object is present on current web page. :
                    #region Action->VerifyObjectPresent
                    case "verifyobjectpresent":
                        verification = _objTestObject.VerifyObjectPresent();
                        break;
                    #endregion
                    //Verify that specified test object is present on current web page. :
                    case "verifyobjectnotpresent":
                        verification = _objTestObject.VerifyObjectNotPresent();
                        break;
                    //Verify test object property. :
                    #region Action-> VerifyObjectProperty
                    case "verifyobjectproperty":
                        property = contentFirst.Trim();
                        propertyValue = contentSecond.Trim();
                        verification = _objTestObject.VerifyObjectProperty(property, propertyValue, _keywordDic);

                        break;
                    #endregion
                    //Verify test object property not present. :
                    #region Action-> VerifyObjectNotPresent
                    case "verifyobjectpropertynot":
                        property = contentFirst.Trim();
                        propertyValue = contentSecond.Trim();
                        verification = _objTestObject.VerifyObjectPropertyNot(property, propertyValue, _keywordDic);
                        break;
                    #endregion
                    //Verify current web page property. :
                    #region Action->VerifyPageProperty
                    case "verifypageproperty":
                        property = _testData.Substring(0, _testData.IndexOf('|')).Trim();
                        propertyValue = _testData.Substring(_testData.IndexOf('|') + 1).Trim();
                        // Replace special characters after trimming off white spaces.
                        property = Utility.ReplaceSpecialCharactersInString(property);
                        propertyValue = Utility.ReplaceSpecialCharactersInString(propertyValue);
                        verification = _objBrowser.VerifyPageProperty(property, propertyValue, _keywordDic);
                        break;
                    #endregion
                    //Verify that specified text is present in web page view source. :
                    #region Action-> VerifyTextInPageSource
                    case "verifytextinpagesource":
                        verification = _objBrowser.VerifyTextInPageSource(contentFirst, _keywordDic);
                        break;
                    #endregion
                    //Verify specified text not in webpage view-source. 
                    #region Action->VerifyTextNotPageSource
                    case "verifytextnotinpagesource":
                        verification = !_objBrowser.VerifyTextInPageSource(contentFirst, _keywordDic);
                        break;
                    #endregion
                    //Verify that current web page is display properly.
                    case "verifypagedisplayed":
                        verification = _objBrowser.VerifyPageDisplayed(_keywordDic);
                        break;
                    //Synchronize object with specified condition.
                    case "syncobject":
                        _objTestObject.WaitForObject(contentFirst, _globalTimeout, _keywordDic, modifier);
                        break;
                    //Verify that object is dispaly properly.
                    case "verifyobjectdisplayed":
                        verification = _objTestObject.VerifyObjectDisplayed();
                        break;
                    //Execute specified script language.    
                    #region Action->ExecuteStatement | ExecuteScript
                    case "executescript":
                    case "executestatement":
                        Property.NoWait = true;
                        string scriptResult = _objTestObject.ExecuteStatement(contentFirst);
                        if (scriptResult.ToLower().Equals("false"))
                            verification = false;
                        else if (!scriptResult.ToLower().Equals("true"))
                            Utility.SetVariable(_testObject, scriptResult);
                        break;
                    #endregion
                    //Retrieve all numbers from text.
                    case "getnumberfromtext":
                        Utility.GetNumberFromText(contentFirst);
                        break;
                    #region Action->SetVariable
                    case "setvariable":
                        string varName = contentFirst.Trim();
                        string varValue = contentSecond.Trim();
                        //ReplaceVariablesInString will find value from $variable
                        varValue = Utility.ReplaceVariablesInString(varValue);
                        Property.StepDescription = "Store " + varValue + " to variable named " + varName;
                        //set variable
                        Utility.SetVariable(varName, varValue);

                        verification = true; //should this step need to be captured in the log file as pass/fail?

                        break;
                    #endregion
                    #region Action->SetParameter
                    case "setparameter":

                        string varName1 = contentFirst.Trim();
                        string varValue1 = contentSecond.Trim();

                        //ReplaceVariablesInString will find value from $variable
                        varValue = Utility.ReplaceVariablesInString(varValue1);
                        Property.StepDescription = "Store " + varValue1 + " to variable named " + varName1;

                        //set parameter
                        Utility.SetParameter(varName1, varValue1);

                        //set variable
                        Utility.SetVariable(varName1, varValue1);

                        verification = true;

                        break;
                    #endregion
                    case "getdatafromdatabase":
                        verification = DBRelatedMethods.GetDataFromDatabase(contentFirst, child);
                        _verificationMessage = Property.Remarks;
                        break;
                    #region Action-> GetUserTesting
                    case "getuserfortesting": //it will work same as method "GetDataFromDatabase". If GetUserForTesting returns false, test case execution will be stopped.
                        try
                        {
                            verification = DBRelatedMethods.GetDataFromDatabase(contentFirst, child);
                            _verificationMessage = Property.Remarks;

                            //Try once again in case this fails
                            if (verification == false)
                            {
                                verification = DBRelatedMethods.GetDataFromDatabase(contentFirst, child);
                                _verificationMessage = Property.Remarks;
                            }

                            if (verification == false) Property.EndExecutionFlag = true;
                        }
                        catch
                        {
                            Property.EndExecutionFlag = true;
                        }
                        break;
                    #endregion
                    case "executedatabasequery":
                        verification = DBRelatedMethods.ExecuteDatabaseQuery(contentFirst, child);
                        _verificationMessage = Property.Remarks;
                        break;
                    case "verifydatabase":
                        verification = DBRelatedMethods.VerifyDatabase(contentFirst, child);
                        _verificationMessage = Property.Remarks;
                        break;
                    #region Action->RequestWebService
                    case "requestwebservice":
                        string requestHeader = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Header"));
                        string responseFormat = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Format"));
                        string requestMethod = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Method"));
                        string requestUrl = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]ServiceURL"));
                        string requestBody = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Body"));

                        Property.StepDescription = "Request web service using method '" + requestMethod + "' and url '" + requestUrl + "'.";
                        verification = WebAPI.RequestWebService(requestHeader, responseFormat, requestMethod, requestUrl, requestBody);
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    case "downloadfile":
                    case "filedownload":
                    #region Action->Download
                    case "download":
                        Property.StepDescription = "Download file from url: " + contentFirst;
                        verification = WebAPI.DownloadFile(contentFirst, contentSecond);
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    #region Action-> FTPFileUpload
                    case "ftpfileupload":
                        Property.StepDescription = "Upload file: " + contentFourth + "to ftp server: " + contentFirst;

                        string hostName = contentFirst;
                        string userName = contentSecond;
                        string password = contentThird;
                        string filePath = contentFourth;
                        string uploadedFileName = contentFifth;

                        string url = string.Empty;

                        if (String.IsNullOrEmpty(uploadedFileName))
                        {
                            url = "ftp://" + hostName + "/" + Path.GetFileName(filePath);
                        }
                        else
                        {
                            url = "ftp://" + hostName + "/" + uploadedFileName;
                        }

                        Uri uploadUrl = new Uri(url);

                        verification = WebAPI.UploadFiletoFtp(filePath, uploadUrl, userName, password);
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    #region Action-> ReadXMLAttribute
                    case "readxmlattribute":
                        Property.StepDescription = "Read value of attribute '" + contentFirst.Trim() +
                            "' and store to run time variable '" + contentFirst.Trim() + "'.";

                        //Pass on node index required as a seconf argument if provided in test data
                        if (!contentSecond.Equals(String.Empty))
                        {
                            try
                            {
                                verification = WebAPI.ReadXmlAttribute(contentFirst.Trim(),
                                                                       Convert.ToInt16(contentSecond.Trim()));
                            }
                            catch (Exception)
                            {
                                verification = WebAPI.ReadXmlAttribute(contentFirst.Trim());
                            }
                        }
                        else
                        {
                            verification = WebAPI.ReadXmlAttribute(contentFirst.Trim());
                        }

                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    case "findxmlattribute":
                    case "locatexmlattribute":
                    #region Action-> GetXMLAttributePosition
                    case "getxmlattributeposition":
                        Property.StepDescription = "Find relative location (index) of attribute '" + contentFirst.Trim() +
                            "' where value of attribute should be '" + contentSecond.Trim() + "'.";

                        verification = WebAPI.ReadXmlAttribute(contentFirst.Trim(), 1, contentSecond.Trim());
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    // New step action added as per the need of G3 API automation testing team  
                    case "verifyxmlattributenotpresent":
                        verification = !WebAPI.ReadXmlAttribute(contentFirst.Trim(), 1, contentSecond.Trim());
                        if (!verification)
                            Property.Remarks = "Value '" + contentSecond.Trim() + "' of attribute '" + contentFirst.Trim() + "' found  in api response. ";
                        else
                            Property.Remarks = "Value '" + contentSecond.Trim() + "' of attribute '" + contentFirst.Trim() + "' was not present in api response. "; ;
                        _verificationMessage = Property.Remarks;
                        break;
                    #region Action-> VerifyXMLAttribute
                    case "verifyxmlattribute":
                        Property.StepDescription = "Verify value of attribute '" +
                            contentFirst.Trim() + "' equals to '" +
                            contentSecond.Trim() + "'.";
                        //Third attribute is considered to be 
                        if (contentThird.Equals(string.Empty))
                        {
                            contentThird = "1";
                        }

                        verification = WebAPI.VerifyXmlAttribute(contentFirst.Trim(), contentSecond.Trim(), int.Parse(contentThird));
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion

                    #region Action->VerifyXMLHeader
                    case "verifyxmlheader":
                        verification = WebAPI.VerifyXmlHeader(contentFirst.Trim(), contentSecond.Trim());
                        if (verification)
                        {
                            Property.Remarks = contentFirst.Trim()  +" "+contentSecond.Trim() + " Found in the Request Header.";
                        } 
                        else
                            Property.Remarks = contentFirst.Trim()  +" "+contentSecond.Trim() + " Not found in the Request Header.";
                                _verificationMessage = Property.Remarks;
                        break;

                    #endregion
                    #region Action-> CountXMLNodes
                    case "countxmlnodes":
                        Property.StepDescription = "Count all instances of attribute '" + contentFirst + "' and store to variable '" + contentFirst.Trim() + "Count'";
                        verification = WebAPI.CountXmlNodes(contentFirst.Trim());
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    #region Action-> VerifyXMLNodeCount
                    case "verifyxmlnodecount":
                        Property.StepDescription = "Verify count of all instances of attribute '" +
                                contentFirst.ToString().Trim() +
                                "' equals to '" + contentSecond.ToString().Trim() + "'.";

                        verification = WebAPI.Verifyxmlnodecount(contentFirst.ToString().Trim(),
                                                                    contentSecond.ToString().Trim());
                        _verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    // this method will split data column, generate random number on the basis of start index and end index and assign to variable
                    // 
                    #region Action-> GenerateRandomNumber
                    case "generaterandomnumber":
                        int startIntValue = int.Parse(Utility.ReplaceVariablesInString(contentFirst.Trim()));
                        int endIntValue = int.Parse(Utility.ReplaceVariablesInString(contentSecond.Trim()));
                        string generateParamName = string.Empty;
                        if (!contentThird.Equals(String.Empty))
                            generateParamName = Utility.ReplaceVariablesInString(contentThird.Trim());
                        else
                            generateParamName = Property.GenerateRandomNumberParamName; //by default harcoded variable

                        int rndVal = Utility.RandomNumber(startIntValue, endIntValue + 1);
                        //set variable
                        Utility.SetVariable(generateParamName, rndVal.ToString());
                        break;
                    #endregion
                    //this method will split data column, generate random number on the basis of start index and end index and assign to variable
                    // 
                    #region Action-> GenerateUniqueString
                    case "generateuniquestring":
                        int startIntValue1 = int.Parse(Utility.ReplaceVariablesInString(contentFirst.Trim()));
                        string endIntValue1 = Utility.ReplaceVariablesInString(contentSecond.Trim());
                        string generateParamName1 = string.Empty;
                        if (string.IsNullOrWhiteSpace(endIntValue1) == false)
                            generateParamName1 = endIntValue1;
                        else
                            generateParamName1 = Property.GenerateRandomStringParamName; //by default harcoded variable

                        string rndVal1 = Utility.RandomString(startIntValue1);

                        //set variable
                        Utility.SetVariable(generateParamName1, rndVal1.ToString());

                        Property.StepDescription = "Generate unique string of length " + startIntValue1;
                        Property.Remarks = "Generated string:= " + rndVal1.ToString();

                        break;
                    #endregion
                    //This action generates a 10 digit unique number where first digit will not be 0 or 1.
                    //Also stores each digit in ten different variables named as: Char1, Char2, ... , Char10
                    #region Action-> GenerateUniqueMobileNo.
                    case "generateuniquemobileno":
                        string uniqueMobile = Utility.GenerateUniqueNumeral(10);
                        // Check if first digit is '1' then replace it with '4'
                        if (uniqueMobile[0] == '0' || uniqueMobile[0] == '1')
                            uniqueMobile = uniqueMobile.Replace(uniqueMobile[0], '4');

                        // Store generated mobile number into the dictionary
                        Utility.SetVariable("MobileNo", uniqueMobile);

                        // Also store each character of mobile number into the dictionary
                        for (int j = 1; j <= uniqueMobile.Length; j++)
                            Utility.SetVariable("Char" + j.ToString(), uniqueMobile[j - 1].ToString());

                        Property.StepDescription = "Generate unique mobile number and store to variable: 'MobileNo'";
                        Property.Remarks = "Generated mobile no:= " + uniqueMobile.ToString();

                        break;
                    #endregion
                    case "mousemove":
                        _objTestObject.MouseMove();
                        break;

                    //Performs click using advanced user interaction API
                    case "mouseclick":
                        _objTestObject.MouseClick();
                        break;

                    //Performs mouseover using advanced user interaction API
                    case "mouseover":
                        _objTestObject.MouseOver();
                        break;
                    case "{ignore}":
                    case "{IGNORE}":
                    case "{SKIP}":
                    case "{skip}":
                        stepStatus = ExecutionStatus.Pass;
                        break;

                    //Builds an advanced user interaction API object, which should later be used by "performAction" keyword
                    //            Multiple addAction statements can be clubbed one after another, and they will be executed in a chain fashion
                    #region Action-> AddAction
                    case "addaction":
                        _objTestObject.AddAction(contentFirst, contentSecond);
                        break;
                    #endregion
                    //Performs actions build using "addAction" keyword using advanced user interaction API
                    #region Action-> PerformAction
                    case "performaction":
                        _objTestObject.PerformAction();
                        break;
                    #endregion
                    #region Action-> VerifySortOrder
                    case "verifysortorder":
                        Property.StepDescription = "Verify sort order with Property " + "'" + contentFirst + "'" + " and sort order " + "'" + contentSecond + "'";
                        verification = _objTestObject.VerifySortOrder(contentFirst, contentSecond);
                        break;
                    #endregion

                    #region Action-> DragAndDrop
                    case "draganddrop":
                        Property.StepDescription = "Drag " + _testObject + " and drop to " + contentFirst;
                        _objTestObject.DragAndDrop(ObjSecondDataRow);
                        break;
                    #endregion
                    #region  Action-> saveContentsoffile
                    case "savecontentstofile":
                        string contents = contentFirst;
                        string filename = "CustomTagsFile.txt";
                        if (!_testObject.Equals(string.Empty))
                        {
                            filename = _testObject;
                        }
                        string filepath = Property.ResultsSourcePath + "\\" + filename;
                        string varname = filename.Split('.')[0].ToString();
                        Utility.Savecontentstofile(filepath, contents);
                        Utility.SetVariable(varname, filepath);
                        Property.Remarks = "Data Content has been stored to file '" + filename + "'.Its location has been stored in a variable '" + varname + "'.";
                        break;
                    #endregion
                    #region UploadFile  PBI : #205
                    case "uploadfile":
                        Property.StepDescription = "Upload file at location " + contentFirst;
                        verification = _objTestObject.UploadFile(contentFirst, contentSecond);
                        if (verification == false)
                        {
                            Property.Remarks = "File not found at location " + contentFirst;
                        }
                        break;

                    #endregion
                    default:
                        {
                            Property.Remarks = _stepAction + ":" + Utility.GetCommonMsgVariable("KRYPTONERRCODE0025");
                            verification = false;
                            stepStatus = ExecutionStatus.Fail;
                        }
                        break;

                }
                if (!stepStatus.Equals(ExecutionStatus.Warning))
                {
                    stepStatus = verification ? ExecutionStatus.Pass : ExecutionStatus.Fail;

                }
            }

            catch (Exception e)
            {
                if (modifier.Contains("recovery"))
                {
                    throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0024"));
                }

                Property.Remarks = _stepAction + ":" + e.Message;
                stepStatus = ExecutionStatus.Fail;

                //Overriding cache errors that appears on saucelabs
                //seen when testing for angieslist
                if (e.Message.IndexOf("getElementTagName execution failed; Element does not exist in cache", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;
                if (e.Message.IndexOf("Element not found in the cache", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Pass;

                //Overriding timeouts to warning instead of failure
                if (e.Message.IndexOf(Exceptions.ERROR_NORESPONSEURL, StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;

                if (e.Message.IndexOf("docElement is null", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;

                //When this message appears, click has been successfull, just the error appears
                if (e.Message.IndexOf("Cannot click on element", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;

                //Overriding click timeout to waring instead of failure
                if (e.Message.IndexOf("Timed out after", StringComparison.OrdinalIgnoreCase) >= 0 || e.Message.IndexOf("Page load is taking longer time.", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;

                //Overriding timeout to warning instead of failure. 
                if (e.Message.IndexOf("win is null", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;
                //Modal dialog present Execption does,t interrupt normal Execution 
                if (e.Message.IndexOf("Modal dialog", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Property.Remarks = string.Empty;
                    stepStatus = ExecutionStatus.Pass;
                }


            }

            Property.Status = stepStatus;
            if (_snapShotOption.ToLower().Equals("always")
                || (_snapShotOption.ToLower().Equals("on page change") && (action.ToLower().Equals("click")
                                                                            || action.ToLower().Equals("openbrowser")
                                                                            || action.ToLower().Equals("closebrowser")
                                                                            || action.ToLower().Equals("navigateurl")
                                                                            || action.ToLower().Equals("goback")
                                                                            || action.ToLower().Equals("refreshbrowser")
                                                                            || action.ToLower().Equals("goforward")
                                                                            || action.ToLower().Equals("switchtorecentbrowser")
                                                                            || action.ToLower().Equals("switchtonewbrowser")
                                                                            || action.ToLower().Equals("doubleclick")
                                                                            || action.ToLower().Equals("submit")
                                                                            ))
                || (stepStatus.Equals(ExecutionStatus.Fail) && !_snapShotOption.ToLower().Equals("never") && !action.Contains("."))
                )
            {
                SaveSnapShots(action);
            }

            if (_keywordDic.ContainsValue("endtestonfailure") && stepStatus.Equals(ExecutionStatus.Fail))
            {
                Property.EndExecutionFlag = true;
            }

            if (_testObject != null && parent != null && Property.StepDescription.Equals(string.Empty))
            {
                if (!_testObject.Equals(string.Empty) && !parent.Equals(string.Empty))
                    Property.StepDescription = _stepAction + " on test object \"" + _testObject + "\" in page \"" + parent + " \".";
                else
                    Property.StepDescription = _stepAction; //For actions in which Parent and TestObject doesn't use like closeAllBrowsers etc.
            }
        }

        private void SaveSnapShots(string stepAction = "")
        {
            try
            {
                var filesName = _objBrowser.GetScreenShot(Property.StepNumber, Property.ResultsSourcePath, stepAction);
                // because string may be empty or null
                if (string.IsNullOrWhiteSpace(filesName) == false) 
                {
                    string[] temp = filesName.Split('|');
                    Property.Attachments = temp[0].Trim();
                    if (temp.Length > 1)
                        Property.HtmlSourceAttachment = temp[1].Trim();
                    else Property.HtmlSourceAttachment = string.Empty;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error in saving the snapShot : {0}", e.Source);
            }
        }
        public static void ResetApp()
        {

        }

    }

}
