/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: DialogHandler.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Method Contain Handle the Browser PopUp and Fill Information if Required.
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Threading;
using ControllerLibrary;
using Common;

namespace TestDriver
{
    class DialogHandler
    {
        Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        private string _objType = string.Empty;
        private readonly string _attributeType = string.Empty;
        private string attribute = string.Empty;
        private string _browserClassName = string.Empty;
        private AutomationElement _dialogElement;
        public DialogHandler()
        {
            //Retreive content from OR row dictionary.
            if (Action.ObjDataRow.Count > 0)
            {
                objDataRow = Action.ObjDataRow;
                _objType = objDataRow[KryptonConstants.OBJ_TYPE];
                _attributeType = objDataRow[KryptonConstants.HOW];
                attribute = objDataRow[KryptonConstants.WHAT];
            }
            else
            {
                objDataRow = Action.ObjDataRow;
                _objType = string.Empty;
                _attributeType = string.Empty;
                attribute = string.Empty;
            }

        }

        private void SetClassNameForBrowser()
        {
            //Working fine for all FF version(Win XP,WIn 7 ).
            if (Utility.GetParameter(Property.BrowserString).Equals("ie"))
                _browserClassName = "IEFrame";
            else if (Utility.GetParameter(Property.BrowserString).ToLower().Equals("firefox"))
            {
                string versionString = Utility.GetFfVersionString();
                if (versionString.Contains("4.0"))
                    _browserClassName = "MozillaWindowClass";
                else
                    _browserClassName = "MozillaUIWindowClass";
            }
        }
        #region Set Automation dialog Element.
        /// <summary>
        /// Set Final element on which actual operation would be performed.
        /// </summary>
        private AutomationElement SetAutomationElement()
        {
            try
            {
                // Set Class Name based on browser.
               

                //Get Root Element. This will be pointing to desktop.
                AutomationElement rootElement = AutomationElement.RootElement;

                //Update root element to point to browser instead of desktop
                

                //Split What field in OR.
                string[] whatContents = Regex.Split(attribute, "//");

                string objectLocator;
                if (whatContents.Length.Equals(1))
                {
                    objectLocator = whatContents[0];
                }
                else
                {
                    //Get Dailog Title
                    var dialogText = whatContents[0].ToLower();

                    //Update root element to point to specific dialog
                    Condition parentCondition = new PropertyCondition(AutomationElement.NameProperty, dialogText, PropertyConditionFlags.IgnoreCase);

                    //Find parent window on desktop
                    rootElement = rootElement.FindFirst(TreeScope.Descendants, parentCondition);

                    //Get Element Locator.
                    objectLocator = whatContents[1];
                }

                _dialogElement = getElement(rootElement, objectLocator);//Get the dialog element.
                return _dialogElement;
            }
            catch (IndexOutOfRangeException e)
            {
                throw e;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion

        /// <summary>
        /// Implicitely wait for Win32 Objects within a specified Time range.
        /// </summary>
        /// <returns>AutomationElement : Win32 Object.</returns>
        private AutomationElement waitAndGetElement()
        {
            double diff = 0;
            double start_time = 0;
            double end_time = 0;
            AutomationElement element = null;
            int maxTime = int.Parse(Utility.GetParameter("ObjectTimeout"));
            try
            {
                for (int seconds = 0; ; seconds++)
                {
                    start_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                    if (seconds >= maxTime || diff >= maxTime)
                    {
                        throw new TimeoutException();
                    }
                    try
                    {
                        if ((element = SetAutomationElement()) != null)
                        {
                            return element;
                        }
                    }
                    catch (IndexOutOfRangeException e)
                    {
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0029"));
                    }
                    catch (Exception e) 
                    {
                        throw e;
                    }
                    Thread.Sleep(1000);
                    end_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                    diff = diff + (end_time - start_time);
                }
            }
            catch (TimeoutException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Get Automation element based on parent element.
        /// </summary>
        /// <param name="parentElement">AutomationElement : parentElement</param>
        /// <param name="value">string : Value based on which condition would be developed.</param>
        /// <returns></returns>
        private AutomationElement getElement(AutomationElement parentElement, string value)
        {
            Condition condition = null;
            AutomationElement returnElement = null;
            switch (_attributeType.ToLower())
            {
                case "name":
                    condition = new PropertyCondition(AutomationElement.NameProperty, value);
                    break;
                case "id":
                    condition = new PropertyCondition(AutomationElement.AutomationIdProperty, value);
                    break;
                case "classname":
                    condition = new PropertyCondition(AutomationElement.ClassNameProperty, value);
                    break;
                default:
                    break;
            }
            returnElement = parentElement.FindFirst(TreeScope.Descendants, condition);
            return returnElement;
        }

        /// <summary>
        /// Click on dialog normal and radio buttons.
        /// </summary>
        public void ClickDialogButton()
        {
            try
            {
                AutomationElement dialogElement = waitAndGetElement();
                if (dialogElement != null)
                {
                    //Click on normal buttons etc
                    if (dialogElement.Current.ControlType.Equals(ControlType.Button) ||
                        dialogElement.Current.ControlType.Equals(ControlType.Text) ||
                        dialogElement.Current.ControlType.Equals(ControlType.Hyperlink)
                        )
                    {
                        InvokePattern pattern = dialogElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        if (pattern != null) pattern.Invoke();
                    }

                    //Click on Radio buttons.
                    else if (dialogElement.Current.ControlType.Equals(ControlType.RadioButton) ||
                             dialogElement.Current.ControlType.Equals(ControlType.TreeItem))
                    {
                        SelectionItemPattern pattern = dialogElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        if (pattern != null) pattern.Select();
                    }
                    else
                    {
                        InvokePattern pattern = dialogElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        if (pattern != null) pattern.Invoke();
                    }
                }
            }
            catch (TimeoutException e)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0030"));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Enter text in dailog Element.
        /// </summary>
        /// <param name="textToEnter">String : Text to enter.</param>
        public void EnterdataInDialog(string textToEnter)
        {
            try
            {
                AutomationElement dialogElement = waitAndGetElement();
                //Enter data to dailog edit fields.
                if (dialogElement != null)
                {
                    if (dialogElement.Current.ControlType.Equals(ControlType.ComboBox) || dialogElement.Current.ControlType.Equals(ControlType.Edit))
                    {
                        ValuePattern pattern = dialogElement.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        if (pattern != null) pattern.SetValue(textToEnter);
                    }
                }
            }
            catch (TimeoutException)
            {
                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0030"));
            }
        }

    }



    class UploadFileOnRemote
    {
        readonly string _autoitExeName;
        readonly string _remoteIp;
        readonly string _filePath;

        public UploadFileOnRemote(string browser,string remoteIp, string filePath)
        {
            _remoteIp = remoteIp;
            var browserName = browser;
            _filePath = filePath;
            if (browserName.ToLower().Equals(KryptonConstants.BROWSER_FIREFOX))
            {             
                _autoitExeName = "FileUploadFF.exe";
            }
            else if (browserName.ToLower().Equals(KryptonConstants.BROWSER_CHROME))
            {
               _autoitExeName = "FileUploadChrome.exe";
            }
            else
            {               
                _autoitExeName = "FileUploadIE.exe";
            }

        }

        public void UploadFileWithAutoIt()
        {
            try
            {
                ServiceAgent service = new ServiceAgent(_remoteIp); //Node ip          
                service.StartExecutableByName(_autoitExeName, _filePath);
                Console.WriteLine(string.Format("Uploading file  {0}  on  {1}  machine..!!", _filePath, _remoteIp));
            }
            catch(Exception ex)
            {
                throw new Exception( string.Format("Coud not Upload File on {0} machine. Actual Error:{1}",_remoteIp,ex.Message));
            }
        }
    }
}

