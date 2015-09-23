/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Krypton.TestEngine.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Action mapping class
*****************************************************************************/

using System;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Common;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Threading;
using OpenQA.Selenium;
using AutoItX3Lib;
using System.Text.RegularExpressions;

namespace TestDriver
{

    public class Action
    {
        [DllImport("AutoItX3.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static public extern int AU3_MouseUp([MarshalAs(UnmanagedType.LPStr)] string Button);
        private string stepAction = string.Empty;
        private string parent = string.Empty;
        private string testObject = string.Empty;
        private string TestData = string.Empty;
        public static string browserVersion = string.Empty;
        private Dictionary<int, string> keywordDic = new Dictionary<int, string>();
        public static Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        public static Dictionary<string, string> objSecondDataRow = new Dictionary<string, string>();
        private Driver.Browser objBrowser = null;
        private Driver.ITestObject objTestObject = null;
        private Driver.RecoveryScenarios objRecovery = null;
        private DialogHandler objHandler = null;
        private IProjectMethods pluginProject = null;
        private string verificationMessage = string.Empty;
        private string snapShotOption = Common.Property.SnapshotOption.ToLower();
        private string globalTimeout = Common.Property.GlobalTimeOut;
        private string debugMode = Common.Property.DebugMode;
        private DataSet datasetRecoverBrowser = new DataSet();
        private DataSet datasetRecoverPopup = new DataSet();
        private DataSet datasetOR = new DataSet();

        // Declared a Win32 function to control other processes' window.
        [DllImport("user32.dll")]
        static extern bool ShowWindow(int hWnd, int nCmdShow);

        /// <summary>
        ///Default constructor initialize TestObject and Browser class
        /// </summary>
        public Action(DataSet recoverPopupData, DataSet recoverBrowserData, DataSet ORData)
        {
            this.datasetRecoverPopup = recoverPopupData;
            this.datasetRecoverBrowser = recoverBrowserData;
            this.datasetOR = ORData;
            string browserName = Common.Utility.GetParameter(Common.Property.BrowserString).ToLower();
           
            objBrowser = new Driver.Browser(Common.Property.ErrorCaptureAs);
            objTestObject = new Driver.TestObject(Utility.GetParameter("ObjectTimeout"));

            string[] availablePlugins = Directory.GetFiles(Common.Property.ApplicationPath, "*ProjectPlugin.dll");
            if (availablePlugins.Length > 0)
            {
                foreach (string availablePlugin in availablePlugins)
                {
                    Assembly asm = Assembly.LoadFrom(availablePlugin);
                    System.Type myType = asm.GetType(asm.GetName().Name + ".MatchProjectPlugin");
                    pluginProject = (IProjectMethods)Activator.CreateInstance(myType);

                }

            }
            //Initialize Object for specified language to create Selenium script.        
            switch (Common.Property.ScriptLanguage.ToLower()) 
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
        /// <summary>
        ///This method will call action from Driver Class. 
        /// </summary>
        /// <param name="action">Step action to perfrom by driver</param>
        /// <param name="parent=">Parent Object</param>
        /// <param name="child">Test object on which operation to be perform</param>
        /// <param name="data">Test Data</param>
        /// <param name="modifier"></param>
        /// <returns></returns>

        public void Do(string action, string parent = null, string child = null, string data = null, string modifier = "")
        {
            snapShotOption = Common.Property.SnapshotOption.ToLower();
            //check for locator directly in test case.
            if (child.Contains("="))
            {
                string how = child.Split('=')[0].ToLower().Trim();
                string what = child.Replace(child.Split('=')[0] + "=", string.Empty);
                objDataRow[KryptonConstants.HOW] = how;
                objDataRow[KryptonConstants.WHAT] = what.Trim();
                objDataRow[KryptonConstants.LOGICAL_NAME] = string.Empty;
                objDataRow[KryptonConstants.OBJ_TYPE] = string.Empty;
                objDataRow[KryptonConstants.MAPPING] = string.Empty;
            }
            stepAction = action;
            testObject = child;
            TestData = data;
            int stindex;
            int endindex;
            //parse modifier
            modifier = modifier.ToLower().Trim();
            int i = 1;
            keywordDic.Clear();//clearing previous keyword.

            string browserDimension = null; //browser dimensions
            for (int v = 0; ; v++)
            {
                if (modifier.Contains("{"))
                {
                    stindex = modifier.IndexOf("{");
                    modifier = modifier.Remove(stindex, 1);
                    endindex = modifier.IndexOf("}");
                    string KeyVariable = modifier.Substring(stindex, (endindex - stindex));
                    if (KeyVariable.ToLower().Contains("windowsize") || KeyVariable.ToLower().Contains("window"))
                        browserDimension = KeyVariable;


                    keywordDic.Add(i, KeyVariable);
                    i++;
                    stindex = modifier.IndexOf("}");
                    modifier = modifier.Remove(stindex, 1);
                }
                else
                {
                    break;
                }
            }

            if (keywordDic.ContainsValue("nowait"))
                Common.Property.NoWait = true;
            else
                Common.Property.NoWait = false;

            string stepStatus = string.Empty;
            bool verification = true;
        
            Common.Utility.driverKeydic = null;
            Common.Utility.driverKeydic = keywordDic;

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
                dataContent = data.Split(Property.SEPERATOR);
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
                    default:
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
                objBrowser.SetObjDataRow(objDataRow);
                objTestObject.SetObjDataRow(objDataRow, stepAction);          
                objHandler = new DialogHandler();
                switch (stepAction.ToLower())
                {
                    //action are performed accoding to below action steps.+
                    #region Action-> settestmode
                    case "settestmode":
                        // Assuming page object is given and Mode is paased as a data field. 
                        Common.Property.NoWait = true;
                        string TestModeVariable = Property.TESTMODE;
                        if (!string.IsNullOrWhiteSpace(contentSecond))
                        {
                            TestModeVariable = Property.TESTMODE + "[" + contentFirst + "]";
                        }

                        //Update test mode only if object could be located
                        if (objTestObject.VerifyObjectPresent())
                        {
                            Utility.SetVariable(TestModeVariable, contentSecond.ToLower().Trim());
                        }

                        //Display test mode on console if debug mode is true
                        if (Utility.GetParameter("debugmode").ToLower().Equals("true"))
                            Property.Remarks = "Current Execution Test Mode =" + "'" + Utility.GetVariable(TestModeVariable) + "'";
                        break;
                    #endregion
                    #region Action-> ShutDownDriver
                    case "shutdowndriver":
                        objBrowser.Shutdown();
                        break;
                    #endregion
                    //Delete all internet temporery file.
                    #region Action-> ClearBrowserCache
                    case "clearbrowsercache":
                        Property.StepDescription = "Clear Browser Cache";
                        BrowserManager.browser.clearCache();
                        break;
                    #endregion
                    //Close most recent tab of browser associated with driver. 
                    #region Action-> CloseBrowser
                    case "closebrowser":
                        Property.StepDescription = "Close Browser";
                        objBrowser.CloseBrowser();
                        break;
                    #endregion
                    //Close All Browsers
                    #region Action-> CloseAllBrowsers
                    case "closeallbrowsers":
                        if (Common.Utility.GetParameter(Common.Property.BrowserString).ToLower() != "android" && Common.Utility.GetParameter(Common.Property.BrowserString).ToLower() != "iphone" && Common.Utility.GetParameter(Common.Property.BrowserString).ToLower() != "selendroid")
                        {
                            Property.StepDescription = "Close all opened browsers.";
                            objBrowser.CloseAllBrowser();
                        }
                        break;
                    #endregion
                    //Fire specified event on test object, eg.Click event.  
                    #region Action-> FireEvent
                    case "fireevent":
                        Common.Property.NoWait = true;
                        Property.StepDescription = "Fire '" + contentFirst + "' event on " + testObject;
                        objTestObject.FireEvent(contentFirst);
                        break;
                    #endregion
                    //Start new instance of driver and open browser with specified url 
                    #region Action->OpenBrowser
                    case "openbrowser":
                        string browserName = Common.Utility.GetParameter(Common.Property.BrowserString).ToLower();
                        bool deleteCookie = !modifier.ToLower().Contains("keepcookies");
                            Property.StepDescription = "Open a new browser and navigate to url '" + contentFirst;
                            string remoteUrl = Common.Property.RemoteUrl;
                            string isRemoteExecution = Common.Property.IsRemoteExecution;
                            string ProfilePath = string.Empty;
                            // firefoxProfilePath parameter determines whether to use Firefox Profile or not.
                            switch (browserName)
                            {
                                case KryptonConstants.BROWSER_CHROME: ProfilePath = Utility.GetVariable("ChromeProfilePath");
                                    break;
                                case KryptonConstants.BROWSER_FIREFOX: ProfilePath = Utility.GetVariable("FirefoxProfilePath");
                                    break;
                                default: break;
                            }

                            //addonsPath parameter determines whether to load firefox addons or not.
                            string addonsPath = Utility.GetParameter("AddonsPath");

                            if (contentFirst.Equals(string.Empty))
                            {
                                contentFirst = Common.Property.ApplicationURL;
                            }
                            
                            Exception openBrowserEx = null;
                            try
                            {
                               
                                objBrowser = Driver.Browser.OpenBrowser(browserName, deleteCookie, contentFirst,
                                                                            isRemoteExecution, remoteUrl, ProfilePath,
                                                                            addonsPath, datasetRecoverPopup, datasetRecoverBrowser, datasetOR, browserDimension);
                            }
                            catch (Exception ex)
                            {
                                openBrowserEx = ex;
                                if(Property.IsRemoteExecution.ToLower().Equals("true"))
                                  throw ex;

                            }

                            objTestObject = new Driver.TestObject(Utility.GetParameter("ObjectTimeout"));
                            objRecovery = new Driver.RecoveryScenarios(datasetRecoverPopup, datasetRecoverBrowser, datasetOR, objTestObject);

                            #region  Region containing JavaScript to maximize window and get Browser version string.
                            string browserVer = string.Empty;
                            Common.Utility.SetVariable("BrowserVersion", browserVersion);
                            try
                            {
                                browserVer = objTestObject.ExecuteStatement("return navigator.userAgent;");
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
                                        browserVer = browserVer.Substring(browserVer.IndexOf("Firefox/") + 8);
                                        if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                            browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                        browserVersion = browserVer;
                                        break;
                                    case "ie":
                                    case "iexplore":
                                    case "internetexplorer":
                                    case "internet explorer":
                                        browserVersion = browserVer.Substring(browserVer.IndexOf("MSIE ") + 5).Split(';')[0];
                                        string keyName = null;
                                        // Read the system registry to get the IE version in case tests are running locally
                                        if (!Utility.GetParameter("RunRemoteExecution").Equals("true", StringComparison.OrdinalIgnoreCase))
                                        {
                                            keyName = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer";
                                           browserVersion = (string)Registry.GetValue(keyName, "svcVersion", "key not found in Registry");
                                            if (browserVersion.Equals("key not found in Registry"))
                                            {
                                                browserVersion = (string)Registry.GetValue(keyName, "Version", "key not found in Registry");
                                                if (browserVersion.Equals("key not found in Registry")) // chek if version key is also not available in registry.
                                                {
                                                    browserVersion = string.Empty;
                                                }
                                                else
                                                    browserVersion = browserVersion.Split('.')[0] + "." + browserVersion.Split('.')[1];
                                            }
                                            else
                                            {
                                                browserVersion = browserVersion.Split('.')[0] + "." + browserVersion.Split('.')[1];
                                            }
                                        }
                                        break;
                                    case KryptonConstants.BROWSER_CHROME:                                    

                                        browserVer = browserVer.Substring(browserVer.IndexOf("Chrome/") + 7).Split(' ')[0];
                                        if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                            browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                        browserVersion = browserVer;
                                        break;
                                    case KryptonConstants.BROWSER_OPERA:
                                        if (browserVer.Contains("Version/"))
                                            browserVer = browserVer.Substring(browserVer.IndexOf("Version/") + 8);
                                        else
                                            browserVer = browserVer.Substring(browserVer.IndexOf("Opera") + 6).Split(' ')[0];
                                        if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                            browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                        browserVersion = browserVer;
                                        break;
                                    case KryptonConstants.BROWSER_SAFARI:
                                        browserVer = browserVer.Substring(browserVer.IndexOf("Version/") + 8).Split(' ')[0];
                                        if (browserVer.IndexOf('.') != browserVer.LastIndexOf('.'))
                                            browserVer = browserVer.Split('.')[0] + '.' + browserVer.Split('.')[1];
                                        browserVersion = browserVer;
                                        break;
                                }
                            }
                            #endregion

                            Common.Utility.SetVariable("BrowserVersion", browserVersion);

                            objBrowser.SetObjDataRow(objDataRow);
                            Utility.SetVariable(Common.Property.BrowserVersion, browserVersion);
                        break;
                    #endregion
                    //Navigate to back on current web page
                    #region Action->GoBack
                    case "goback":
                        objBrowser.GoBack();
                        Property.StepDescription = "Go backwards in the browser";
                        break;
                    #endregion
                    //Navigate to forward on current web page. :
                    #region Action->GoForward
                    case "goforward":
                        Property.StepDescription = "Go forward in the browser";
                        objBrowser.GoForward();
                        break;
                    #endregion
                    //Refresh current web page. :
                    #region Action->RefreshBrowser
                    case "refreshbrowser":
                        Property.StepDescription = "Refresh browser";
                        objBrowser.Refresh();
                        break;
                    #endregion
                    case "switchtorecentbrowser":
                    case "switchtonewbrowser":
                        Property.StepDescription = "Set focus to most recently opened window";
                        objBrowser.setBrowserFocus();
                        break;

                    //Navigate to new specified url in current web page.
                    #region Action->NavigateURL
                    case "navigateurl":
                        Property.StepDescription = "Navigate to url '" + contentFirst + "' in currently opened browser";
                        try
                        {
                            if (!contentFirst.StartsWith(@"http://"))
                                contentFirst = @"http://" + contentFirst;
                        }
                        catch { }
                        objBrowser.NavigationUrl(contentFirst);
                        break;
                    #endregion
                    //Clear existing data from test object. 
                    case "clear":
                    case "cleartext":          
                        Property.StepDescription = "Clear object '" + testObject + "'";
                        objTestObject.ClearText();
                        break;
                    //Check radio button and check box. :
                    case "check":
                        Property.StepDescription = "Check '" + contentFirst + "' checkbox '" + testObject + "'";
                        objTestObject.Check(contentFirst);
                        break;
                    //Uncheck radio button and check box. :
                    case "uncheck":
                        Property.StepDescription = "Uncheck checkbox '" + testObject + "'";
                        objTestObject.UnCheck();
                        break;
                    case "checkmultiple":
                        Property.StepDescription = "Check multiple checkboxes of '" + testObject + "'";
                        string[] data_Content = data.Split(Property.SEPERATOR);
                        objTestObject.checkMultiple(data_Content);
                        break;
                    case "uncheckmultiple":
                        Property.StepDescription = "Uncheck multiple checkboxes of '" + testObject + "'";
                        string[] dataC = data.Split(Property.SEPERATOR);
                        objTestObject.uncheckMultiple(dataC);
                        break;
                    case "swipeobject":
                        Property.StepDescription = "swipe object in " + contentFirst + " direction " + testObject;
                        Utility.SetVariable(testObject, contentFirst);
                        break;
                    //Perform click action on associated test object. :
                    #region Action->Click
                    case "click":
                        Property.StepDescription = "Click on '" + testObject + "'";
                                               
                        DateTime dtbefore = new DateTime();
                        if (objDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && objDataRow[KryptonConstants.OBJ_TYPE].ToLower().Contains("winbutton"))
                        {

                            dtbefore = DateTime.Now;
                            if (objDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && objDataRow[KryptonConstants.OBJ_TYPE].ToLower().Contains("winbutton") && Property.IsRemoteExecution.ToLower().Equals("true"))
                            {
                                // do nothing
                            }
                            else
                            {
                                var Autoit = new AutoItX3();
                                string[] windowDetail = Regex.Split(objDataRow[KryptonConstants.WHAT], "//");

                                Autoit.WinActivate(windowDetail[0], string.Empty);

                                Autoit.ControlClick(windowDetail[0], string.Empty, windowDetail[1]);
                            }
                        }
                        else
                        {
                            objTestObject.Click(keywordDic, data);
                        }
                        try
                        {
                        }
                        catch (Exception e)
                        {
                            if (e.Message.Contains("Popup_Text") && !(e.Message.ToLower().Contains("false")))
                                Property.Remarks = Property.Remarks + "  " + e.Message;
                        }
                        break;

                    #endregion
                    case "doubleclick":
                        Property.StepDescription = "Double click on '" + testObject + "'";
                        objTestObject.DoubleClick();
                        break;
                    #endregion
                    //Enter unique data in test object. :
                    #region Action-> EnterUniqueData
                    case "enteruniquedata":
                        Property.StepDescription = "Enter unique string of " + contentFirst + " characters in " + testObject;
                        int length = System.Convert.ToInt16(contentFirst);
                        string strUnique = Utility.GenerateUniqueString(length);//passing length value if mentioned in test case.
                        Utility.SetVariable(testObject, strUnique);
                        objTestObject.SendKeys(strUnique);
                        break;
                    #endregion
                    //Enter specified data in test object. :
                    case "enterdata":
                    case "type":

                    #region Action->TypeString
                    case "typestring":
                        if(contentFirst.Equals("ON", StringComparison.CurrentCultureIgnoreCase))
                            stepAction="Check";
                        if(contentFirst.Equals("OFF", StringComparison.CurrentCultureIgnoreCase))
                             stepAction="Unchek";

                            Property.StepDescription = "Enter text " + contentFirst + " in " + testObject;
                        
                        Utility.SetVariable(testObject, contentFirst);//implicitely set key/Value to runtimedic dictionary before enter any data.
                        if (objDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && objDataRow[KryptonConstants.OBJ_TYPE].ToLower().Equals("winedit"))
                        {
                            if (objDataRow.ContainsKey(KryptonConstants.OBJ_TYPE) && objDataRow[KryptonConstants.OBJ_TYPE].ToLower().Contains("winedit") && Property.IsRemoteExecution.ToLower().Equals("true"))
                            {
                                UploadFileOnRemote oUploadFileOnRemote = new UploadFileOnRemote(Utility.GetParameter("browser"), Property.RemoteMachineIP, contentFirst.Trim());
                                oUploadFileOnRemote.UploadFileWithAutoIt();
                            }
                            else
                            objHandler.enterdataInDialog(contentFirst);
                        }
                        else
                        {
                            objTestObject.SendKeys(contentFirst);
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
                        Property.StepDescription = "Press key " + contentFirst + " on " + testObject;
                        objTestObject.KeyPress(contentFirst);
                        break;
                    #endregion
                    //Adding submit in case this helps when clicks are missing.
                    #region Action->Submit
                    case "submit":
                        Property.StepDescription = "Submit " + testObject;
                        objTestObject.Submit();
                        break;
                    #endregion
                    //Perform wait action for specified time duration.
                    case "pause":
                    #region Action->Wait
                    case "wait":
                        Property.StepDescription = "Pause execution for " + contentFirst + " seconds";
                        browserName = Common.Utility.GetParameter(Common.Property.BrowserString).ToLower();
                        objBrowser.Wait(contentFirst);
                        break;
                    #endregion
                    //Select option form List. 
                    #region Action->SelectItem
                    case "selectitem":
                        Property.StepDescription = "Select '" + contentFirst + "' from " + testObject;
                        Utility.SetVariable(testObject, contentFirst);
                        objTestObject.SelectItem(dataContent);
                        break;
                    #endregion
                    //Select multiple options form List.
                    case "selectmultipleitems":
                    case "selectmultipleitem":
                    #region Action->SelectItems
                    case "selectitems":
                        Property.StepDescription = "Select multiple items '" + data + "' from " + testObject;
                        Utility.SetVariable(testObject, data);
                        objTestObject.SelectItem(dataContent, true);
                        break;
                    #endregion
                    //Select option from List on the besis of index. 
                    #region Action-> SelectItemByIndex
                    case "selectitembyindex":
                        Property.StepDescription = "Select " + contentFirst + "th item from " + testObject;
                        string optionValue = objTestObject.SelectItemByIndex(contentFirst);
                        Utility.SetVariable(testObject, optionValue);
                        break;
                    #endregion
                    //Wait for a specified condition to be happened on test object. 
                    #region Action->WaitForObject
                    case "waitforobject":
                        Property.StepDescription = "Wait until " + testObject + " becomes available";
                        string actualwaitTime = objTestObject.WaitForObject(contentFirst, globalTimeout, keywordDic, modifier.ToLower());
                        break;
                    #endregion
                    //Wait for a specified condition to be happened on test object. 
                    #region Action->WaitForObjectNotPresent
                    case "waitforobjectnotpresent":
                        Property.StepDescription = "Wait until " + testObject + " disappears";
                        objTestObject.WaitForObjectNotPresent(contentFirst, globalTimeout, keywordDic);
                        break;
                    #endregion
                    //Wait for specified property to enable. :
                    #region Action->WaitForProperty
                    case "waitforproperty":
                        Property.StepDescription = "Wait for " + testObject + " to achieve value '" + contentSecond +
                                                    "' for '" + contentFirst + "' property";
                        objTestObject.WaitForObjectProperty(contentFirst, contentSecond, globalTimeout, keywordDic);
                        break;
                    #endregion
                    //Accept alert :
                    #region Action->AcceptAlert
                    case "acceptalert":
                        Property.StepDescription = "Accept alert";
                        objBrowser.AlertAccept();
                        break;
                    #endregion
                    #region Action->DismissAlert
                    case "dismissalert":
                        objBrowser.AlertDismiss();
                        break;
                    #endregion
                    //Verify text on alert page. :
                    #region Action->VerifyAlertText
                    case "verifyalerttext":
                        Property.StepDescription = "Verify alert contains text '" + contentFirst + "'";
                        verification = objBrowser.VerifyAlertText(contentFirst, keywordDic);
                        break;
                    #endregion
                    //Get value of specified property of test object. :
                    case "getattribute":
                    #region Action->GetObjectProperty
                    case "getobjectproperty":
                        string property = contentFirst.Trim();
                        string propertyVariable = testObject + "." + property;
                        if (!contentSecond.Equals(String.Empty))
                            propertyVariable = contentSecond.Trim();
                        string propertyValue = objTestObject.GetObjectProperty(property);
                        Utility.SetVariable(propertyVariable, propertyValue);
                        break;
                    #endregion
                    //set property present in DOM
                    #region Action->SetAttribute
                    case "setattribute":

                        property = contentFirst.Trim();
                        propertyValue = contentSecond.Trim();
                        objTestObject.SetAttribute(property, propertyValue);
                        break;
                    #endregion
                    //Get value of specified property of current web page. :
                    
                    #region Action->GetPageProperty
                    case "getpageproperty":
                        propertyValue = objBrowser.GetPageProperty(contentFirst.Trim());
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
                        if (objDataRow.Count == 0)
                        {
                            verification = objBrowser.VerifyTextPresentOnPage(contentFirst, keywordDic);
                        }
                        else
                        {
                            verification = objTestObject.VerifyObjectProperty("text", contentFirst, keywordDic);
                        }
                        break;
                    #endregion
                    //Verify that specified text is present on current web page or test object. :

                    case "verifytextnotcontained":
                    #region Action->VerifyTextNotPresent
                    case "verifytextnotpresent":

                        if (objDataRow[KryptonConstants.HOW].Equals(string.Empty) || objDataRow[KryptonConstants.WHAT].Equals(string.Empty) || objDataRow[KryptonConstants.HOW] == null || objDataRow[KryptonConstants.WHAT] == null || objDataRow[KryptonConstants.HOW].ToLower().Equals("url"))
                        {
                            verification = !objBrowser.VerifyTextPresentOnPage(contentFirst, keywordDic);
                        }
                        else
                        {
                            verification = objTestObject.VerifyObjectPropertyNot("text", contentFirst, keywordDic);
                        }
                        break;
                    #endregion
                    //Verify that specified text is present on current web page. :
                    #region Action->VerifyTextOnPage
                    case "verifytextonpage":
                        verification = objBrowser.VerifyTextPresentOnPage(contentFirst, keywordDic);
                        if (verification)
                            Property.StepDescription = "The Text : " + contentFirst + " was present on the page";
                        else
                            Property.StepDescription = "The Text : " + contentFirst + " was NOT present on the page";
                        break;
                    #endregion
                    //Verify that specified text is not present on current web page. :
                    #region Action-> VerifyTextNotOnPage
                    case "verifytextnotonpage":
                        verification = !objBrowser.VerifyTextPresentOnPage(contentFirst, keywordDic);
                        if (verification)
                        {
                            stepStatus = ExecutionStatus.Pass;
                            Common.Property.Remarks = "Text  : \"" + contentFirst + "\" is not found on current Page.";
                        }
                        else
                        {
                            stepStatus = ExecutionStatus.Fail;
                            Common.Property.Remarks = "Text  : \"" + contentFirst + "\" is found on current Page.";
                        }
                        break;
                    #endregion
                    //Verify that List option is persent in focused List. :
                    case "verifylistitempresent":
                        verification = objTestObject.VerifyListItemPresent(contentFirst);
                        break;
                    //Verify that List option is not persent in focused List. :
                    case "verifylistitemnotpresent":
                        verification = objTestObject.VerifyListItemNotPresent(contentFirst);
                        break;
                    //Verify that specified test object is present on current web page. :
                    #region Action->VerifyObjectPresent
                    case "verifyobjectpresent":
                        verification = objTestObject.VerifyObjectPresent();
                        break;
                    #endregion
                    //Verify that specified test object is present on current web page. :
                    case "verifyobjectnotpresent":
                        verification = objTestObject.VerifyObjectNotPresent();
                        break;
                    //Verify test object property. :
                    #region Action-> VerifyObjectProperty
                    case "verifyobjectproperty":
                        property = contentFirst.Trim();
                        propertyValue = contentSecond.Trim();
                        verification = objTestObject.VerifyObjectProperty(property, propertyValue, keywordDic);
                                                                   
                        break;
                    #endregion
                    //Verify test object property not present. :
                    #region Action-> VerifyObjectNotPresent
                    case "verifyobjectpropertynot":
                        property = contentFirst.Trim();
                        propertyValue = contentSecond.Trim();
                        verification = objTestObject.VerifyObjectPropertyNot(property, propertyValue, keywordDic);
                        break;
                    #endregion
                    //Verify current web page property. :
                    #region Action->VerifyPageProperty
                    case "verifypageproperty":
                        property = TestData.Substring(0, TestData.IndexOf('|')).Trim();
                        propertyValue = TestData.Substring(TestData.IndexOf('|') + 1).Trim();
                        // Replace special characters after trimming off white spaces.
                        property = Utility.ReplaceSpecialCharactersInString(property);
                        propertyValue = Utility.ReplaceSpecialCharactersInString(propertyValue);
                        verification = objBrowser.VerifyPageProperty(property, propertyValue, keywordDic);
                        break;
                    #endregion
                    //Verify that specified text is present in web page view source. :
                    #region Action-> VerifyTextInPageSource
                    case "verifytextinpagesource":
                        verification = objBrowser.VerifyTextInPageSource(contentFirst, keywordDic);
                        break;
                    #endregion
                    //Verify specified text not in webpage view-source. 
                    #region Action->VerifyTextNotPageSource
                    case "verifytextnotinpagesource":
                        verification = !objBrowser.VerifyTextInPageSource(contentFirst, keywordDic);
                        break;
                    #endregion
                    //Verify that current web page is display properly.
                    case "verifypagedisplayed":
                        verification = objBrowser.VerifyPageDisplayed(keywordDic);
                        break;
                    //Synchronize object with specified condition.
                    case "syncobject":
                        objTestObject.WaitForObject(contentFirst, globalTimeout, keywordDic, modifier);
                        break;
                    //Verify that object is dispaly properly.
                    case "verifyobjectdisplayed":
                        verification = objTestObject.VerifyObjectDisplayed();
                        break;
                    //Execute specified script language.    
                    #region Action->ExecuteStatement | ExecuteScript
                    case "executescript":
                    case "executestatement":
                        Common.Property.NoWait = true;
                        string scriptResult = objTestObject.ExecuteStatement(contentFirst);
                        if (scriptResult.ToLower().Equals("false"))
                            verification = false;
                        else if (!scriptResult.ToLower().Equals("true"))                           
                            Utility.SetVariable(testObject, scriptResult);
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
                        verificationMessage = Property.Remarks;
                        break;
                    #region Action-> GetUserTesting
                    case "getuserfortesting": //it will work same as method "GetDataFromDatabase". If GetUserForTesting returns false, test case execution will be stopped.
                        try
                        {
                            verification = DBRelatedMethods.GetDataFromDatabase(contentFirst, child);
                            verificationMessage = Property.Remarks;

                            //Try once again in case this fails
                            if (verification == false)
                            {
                                verification = DBRelatedMethods.GetDataFromDatabase(contentFirst, child);
                                verificationMessage = Property.Remarks;
                            }

                            if (verification == false) Common.Property.EndExecutionFlag = true;
                        }
                        catch
                        {
                            Common.Property.EndExecutionFlag = true;
                        }
                        break;
                    #endregion
                    case "executedatabasequery":
                        verification = DBRelatedMethods.ExecuteDatabaseQuery(contentFirst, child);
                        verificationMessage = Property.Remarks;
                        break;
                    case "verifydatabase":
                        verification = DBRelatedMethods.verifyDatabase(contentFirst, child);
                        verificationMessage = Property.Remarks;
                        break;
                    #region Action->RequestWebService
                    case "requestwebservice":
                        string requestHeader = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Header"));
                        string responseFormat = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Format"));
                        string requestMethod = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Method"));
                        string requestUrl = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]ServiceURL"));
                        string requestBody = Utility.ReplaceVariablesInString(Utility.GetVariable("[TD]Body"));

                        Property.StepDescription = "Request web service using method '" + requestMethod + "' and url '" + requestUrl + "'.";
                        verification = WebAPI.requestWebService(requestHeader, responseFormat, requestMethod, requestUrl, requestBody);
                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    case "downloadfile":
                    case "filedownload":
                    #region Action->Download
                    case "download":
                        Property.StepDescription = "Download file from url: " + contentFirst;
                        verification = WebAPI.downloadFile(contentFirst, contentSecond);
                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    #region Action-> FTPFileUpload
                    case "ftpfileupload":
                        Property.StepDescription = "Upload file: " + contentFourth + "to ftp server: " + contentFirst;

                        string HostName = contentFirst;
                        string UserName = contentSecond;
                        string Password = contentThird;
                        string FilePath = contentFourth;
                        string UploadedFileName = contentFifth;

                        string url = string.Empty;

                        if (String.IsNullOrEmpty(UploadedFileName))
                        {
                            url = "ftp://" + HostName + "/" + Path.GetFileName(FilePath);
                        }
                        else
                        {
                            url = "ftp://" + HostName + "/" + UploadedFileName;
                        }

                        Uri uploadUrl = new Uri(url);

                        verification = WebAPI.uploadFiletoFTP(FilePath, uploadUrl, UserName, Password);
                        verificationMessage = Property.Remarks;
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
                                verification = WebAPI.readXmlAttribute(contentFirst.Trim(),
                                                                       Convert.ToInt16(contentSecond.Trim()));
                            }
                            catch (Exception)
                            {
                                verification = WebAPI.readXmlAttribute(contentFirst.Trim());
                            }
                        }
                        else
                        {
                            verification = WebAPI.readXmlAttribute(contentFirst.Trim());
                        }

                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    case "findxmlattribute":
                    case "locatexmlattribute":
                    #region Action-> GetXMLAttributePosition
                    case "getxmlattributeposition":
                        Property.StepDescription = "Find relative location (index) of attribute '" + contentFirst.Trim() +
                            "' where value of attribute should be '" + contentSecond.Trim() + "'.";

                        verification = WebAPI.readXmlAttribute(contentFirst.Trim(), 1, contentSecond.Trim());
                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    //: new step action added as per the need of G3 API automation testing team  
                    case "verifyxmlattributenotpresent":
                        verification = !WebAPI.readXmlAttribute(contentFirst.Trim(), 1, contentSecond.Trim());
                        if (!verification)
                            Property.Remarks = "Value '" + contentSecond.Trim() + "' of attribute '" + contentFirst.Trim() + "' found  in api response. ";
                        else
                            Property.Remarks = "Value '" + contentSecond.Trim() + "' of attribute '" + contentFirst.Trim() + "' was not present in api response. "; ;
                        verificationMessage = Property.Remarks;
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

                        verification = WebAPI.verifyXmlAttribute(contentFirst.Trim(), contentSecond.Trim(), int.Parse(contentThird));
                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    #region Action-> CountXMLNodes
                    case "countxmlnodes":
                        Property.StepDescription = "Count all instances of attribute '" + contentFirst + "' and store to variable '" + contentFirst.Trim() + "Count'";
                        verification = WebAPI.CountXMLNodes(contentFirst.Trim());
                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    #region Action-> VerifyXMLNodeCount
                    case "verifyxmlnodecount":
                        Property.StepDescription = "Verify count of all instances of attribute '" +
                                contentFirst.ToString().Trim() +
                                "' equals to '" + contentSecond.ToString().Trim() + "'.";

                        verification = WebAPI.verifyxmlnodecount(contentFirst.ToString().Trim(),
                                                                    contentSecond.ToString().Trim());
                        verificationMessage = Property.Remarks;
                        break;
                    #endregion
                    // this method will split data column, generate random number on the basis of start index and end index and assign to variable
                    // 
                    #region Action-> GenerateRandomNumber
                    case "generaterandomnumber":
                        int startIntValue = int.Parse(Common.Utility.ReplaceVariablesInString(contentFirst.Trim()));
                        int endIntValue = int.Parse(Common.Utility.ReplaceVariablesInString(contentSecond.Trim()));
                        string generateParamName = string.Empty;
                        if (!contentThird.Equals(String.Empty))
                            generateParamName = Common.Utility.ReplaceVariablesInString(contentThird.Trim());
                        else
                            generateParamName = Property.GenerateRandomNumberParamName; //by default harcoded variable

                        int rndVal = Utility.RandomNumber(startIntValue, endIntValue + 1);

                        //set variable
                        Common.Utility.SetVariable(generateParamName, rndVal.ToString());

                        break;
                    #endregion
                    //this method will split data column, generate random number on the basis of start index and end index and assign to variable
                    // 
                    #region Action-> GenerateUniqueString
                    case "generateuniquestring":
                        int startIntValue1 = int.Parse(Common.Utility.ReplaceVariablesInString(contentFirst.Trim()));
                        string endIntValue1 = Common.Utility.ReplaceVariablesInString(contentSecond.Trim());
                        string generateParamName1 = string.Empty;
                        if (string.IsNullOrWhiteSpace(endIntValue1) == false)
                            generateParamName1 = endIntValue1;
                        else
                            generateParamName1 = Property.GenerateRandomStringParamName; //by default harcoded variable

                        string rndVal1 = Utility.RandomString(startIntValue1);

                        //set variable
                        Common.Utility.SetVariable(generateParamName1, rndVal1.ToString());

                        Common.Property.StepDescription = "Generate unique string of length " + startIntValue1;
                        Common.Property.Remarks = "Generated string:= " + rndVal1.ToString();

                        break;
                    #endregion
                    //This action generates a 10 digit unique number where first digit will not be 0 or 1.
                    //Also stores each digit in ten different variables named as: Char1, Char2, ... , Char10
                    #region Action-> GenerateUniqueMobileNo.
                    case "generateuniquemobileno":
                        string uniqueMobile = Common.Utility.GenerateUniqueNumeral(10);
                        // Check if first digit is '1' then replace it with '4'
                        if (uniqueMobile[0] == '0' || uniqueMobile[0] == '1')
                            uniqueMobile = uniqueMobile.Replace(uniqueMobile[0], '4');

                        // Store generated mobile number into the dictionary
                        Utility.SetVariable("MobileNo", uniqueMobile);

                        // Also store each character of mobile number into the dictionary
                        for (int j = 1; j <= uniqueMobile.Length; j++)
                            Utility.SetVariable("Char" + j.ToString(), uniqueMobile[j - 1].ToString());

                        Common.Property.StepDescription = "Generate unique mobile number and store to variable: 'MobileNo'";
                        Common.Property.Remarks = "Generated mobile no:= " + uniqueMobile.ToString();

                        break;
                    #endregion
                    case "mousemove":
                        objTestObject.mouseMove();
                        break;

                    //Performs click using advanced user interaction API
                    case "mouseclick":
                        objTestObject.mouseClick();
                        break;

                    //Performs mouseover using advanced user interaction API
                    case "mouseover":
                        objTestObject.mouseOver();
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
                        objTestObject.addAction(contentFirst, contentSecond);
                        break;
                    #endregion
                    //Performs actions build using "addAction" keyword using advanced user interaction API
                    #region Action-> PerformAction
                    case "performaction":
                        objTestObject.performAction();
                        break;
                    #endregion
                    #region Action-> VerifySortOrder
                    case "verifysortorder":
                        Property.StepDescription = "Verify sort order with Property " + "'" + contentFirst + "'" + " and sort order " + "'" + contentSecond + "'";
                        verification = objTestObject.verifySortOrder(contentFirst, contentSecond);
                        break;
                    #endregion
                 
                    #region Action-> DragAndDrop
                    case "draganddrop":
                        Property.StepDescription = "Drag " + testObject + " and drop to " + contentFirst;
                        objTestObject.DragAndDrop(objSecondDataRow);
                        break;
                    #endregion
                    #region  Action-> saveContentsoffile
                    case "savecontentstofile":
                        string contents = contentFirst;
                        string filename = "CustomTagsFile.txt";
                        if (!testObject.Equals(string.Empty))
                        {
                            filename = testObject;
                        }
                        string filepath = Common.Property.ResultsSourcePath + "\\" + filename;
                        string varname = filename.Split('.')[0].ToString();
                        Utility.Savecontentstofile(filepath, contents);
                        Utility.SetVariable(varname, filepath);
                        Common.Property.Remarks = "Data Content has been stored to file '" + filename + "'.Its location has been stored in a variable '" + varname + "'.";
                        break;
                    #endregion
                    #region UploadFile  PBI : #205
                    case "uploadfile":
                        Property.StepDescription = "Upload file at location " + contentFirst;
                        verification = objTestObject.UploadFile(contentFirst,contentSecond);
                        if(verification==false)
                         {
                            Property.Remarks = "File not found at location " + contentFirst; 
                         }
                        break;
                  
                    #endregion
                    default: {
                            Common.Property.Remarks = stepAction + ":" + Utility.GetCommonMsgVariable("KRYPTONERRCODE0025");
                            verification = false;
                            stepStatus = ExecutionStatus.Fail;
                        }
                        break;

                }
                if (!stepStatus.Equals(ExecutionStatus.Warning))
                {
                    if (verification)
                    {
                        stepStatus = ExecutionStatus.Pass;
                    }
                    else
                    {
                        stepStatus = ExecutionStatus.Fail;
                    }

                }
            }

            catch (Exception e)
            {
                if (modifier.Contains("recovery"))
                {
                    throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0024"));
                }

                Property.Remarks = stepAction + ":" + e.Message;
                stepStatus = ExecutionStatus.Fail;

                //Overriding cache errors that appears on saucelabs
                //seen when testing for angieslist
                if (e.Message.IndexOf("getElementTagName execution failed; Element does not exist in cache", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Warning;

                //as per the discussion on 14th June 11
                if (e.Message.IndexOf("Element not found in the cache", StringComparison.OrdinalIgnoreCase) >= 0)
                    stepStatus = ExecutionStatus.Pass;

                //Overriding timeouts to warning instead of failure
                if (e.Message.IndexOf(exceptions.ERROR_NORESPONSEURL, StringComparison.OrdinalIgnoreCase) >= 0)
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

            if (snapShotOption.ToLower().Equals("always")
                || (snapShotOption.ToLower().Equals("on page change") && (action.ToLower().Equals("click")
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
                || (stepStatus.Equals(ExecutionStatus.Fail) && !snapShotOption.ToLower().Equals("never") && !action.Contains("."))
                )
            {
                this.SaveSnapShots(action);
            }

            if (keywordDic.ContainsValue("endtestonfailure") && stepStatus.Equals(ExecutionStatus.Fail))
            {
                Property.EndExecutionFlag = true;
            }

            if (testObject != null && parent != null && Property.StepDescription.Equals(string.Empty))
            {
                if (!testObject.Equals(string.Empty) && !parent.Equals(string.Empty))
                    Property.StepDescription = stepAction + " on test object \"" + testObject + "\" in page \"" + parent + " \".";
                else
                    Property.StepDescription = stepAction; //For actions in which Parent and TestObject doesn't use like closeAllBrowsers etc.
            }
        }

        private void SaveSnapShots(string stepAction = "")
        {
            try
            {
                string filesName = string.Empty;
                filesName = objBrowser.GetScreenShot(Property.StepNumber, Property.ResultsSourcePath, stepAction);
                if (string.IsNullOrWhiteSpace(filesName) == false) // because string may be empty or null
                {
                    string[] temp = filesName.Split('|');
                    Property.Attachments = temp[0].Trim();
                    if (temp.Length > 1)
                        Property.HtmlSourceAttachment = temp[1].Trim();
                    else Property.HtmlSourceAttachment = string.Empty;
                }
            }
            catch (Exception)
            {

            }
        }
        public static void resetApp()
        {
           
        }

    }

}
