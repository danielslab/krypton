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
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Threading;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium;
using ControllerLibrary;
using Common;

namespace TestDriver
{
    class DialogHandler
    {
        Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        private string objType = string.Empty;
        private string attributeType = string.Empty;
        private string attribute = string.Empty;
        private string browserClassName = string.Empty;
        private AutomationElement dialogElement;
        public DialogHandler()
        {
            //Retreive content from OR row dictionary.
            if (Action.objDataRow.Count > 0)
            {
                this.objDataRow = Action.objDataRow;
                objType = this.objDataRow[KryptonConstants.OBJ_TYPE];
                attributeType = this.objDataRow[KryptonConstants.HOW];
                attribute = this.objDataRow[KryptonConstants.WHAT];
            }
            else
            {
                this.objDataRow = Action.objDataRow;
                objType = string.Empty;
                attributeType = string.Empty;
                attribute = string.Empty;
            }

        }

        private void setClassNameForBrowser()
        {
            //Working fine for all FF version(Win XP,WIn 7 ).
            if (Common.Utility.GetParameter(Common.Property.BrowserString).Equals("ie"))
                browserClassName = "IEFrame";
            else if (Common.Utility.GetParameter(Common.Property.BrowserString).ToLower().Equals("firefox"))
            {
                string versionString = Common.Utility.getFFVersionString();
                if (versionString.Contains("4.0"))
                    browserClassName = "MozillaWindowClass";
                else
                    browserClassName = "MozillaUIWindowClass";
            }
        }
        #region Set Automation dialog Element.
        /// <summary>
        /// Set Final element on which actual operation would be performed.
        /// </summary>
        private AutomationElement setAutomationElement()
        {
            try
            {
                // Set Class Name based on browser.
               

                //Get Root Element. This will be pointing to desktop.
                AutomationElement rootElement = AutomationElement.RootElement;

                //Update root element to point to browser instead of desktop
                

                //Split What field in OR.
                string[] whatContents = Regex.Split(attribute, "//");
                string dialogText = string.Empty;
                string objectLocator = string.Empty;

                if (whatContents.Length.Equals(1))
                {
                    objectLocator = whatContents[0];
                }
                else
                {
                    //Get Dailog Title
                    dialogText = whatContents[0].ToLower();

                    //Update root element to point to specific dialog
                    Condition parentCondition = new PropertyCondition(AutomationElement.NameProperty, dialogText, PropertyConditionFlags.IgnoreCase);

                    //Find parent window on desktop
                    rootElement = rootElement.FindFirst(TreeScope.Descendants, parentCondition);

                    //Get Element Locator.
                    objectLocator = whatContents[1];
                }

                dialogElement = getElement(rootElement, objectLocator);//Get the dialog element.
                return dialogElement;
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
            int maxTime = int.Parse(Common.Utility.GetParameter("ObjectTimeout"));
            try
            {
                for (int seconds = 0; ; seconds++)
                {
                    start_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                    if (seconds >= maxTime || diff >= maxTime)
                    {
                        throw new System.TimeoutException();
                    }
                    try
                    {
                        if ((element = setAutomationElement()) != null)
                        {
                            return element;
                        }
                    }
                    catch (IndexOutOfRangeException e)
                    {
                        throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0029"));
                    }
                    catch (Exception) { }
                    Thread.Sleep(1000);
                    end_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                    diff = diff + (end_time - start_time);
                }
            }
            catch (System.TimeoutException e)
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
            switch (attributeType.ToLower())
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
        public void clickDialogButton()
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
                        pattern.Invoke();

                    }

                    //Click on Radio buttons.
                    else if (dialogElement.Current.ControlType.Equals(ControlType.RadioButton) ||
                             dialogElement.Current.ControlType.Equals(ControlType.TreeItem))
                    {
                        SelectionItemPattern pattern = dialogElement.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                        pattern.Select();
                    }
                    else
                    {
                        InvokePattern pattern = dialogElement.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        pattern.Invoke();

                    }
                }
            }
            catch (TimeoutException e)
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0030"));
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
        public void enterdataInDialog(string textToEnter)
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
                        pattern.SetValue(textToEnter);
                    }
                }
            }
            catch (TimeoutException)
            {
                throw new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0030"));
            }
            catch (Exception)
            {
                throw;
            }
        }

    }



    class UploadFileOnRemote
    {
        string _browserName;
        string _AutoitExeName = string.Empty;
        string _RemoteIp = string.Empty;
        string _FilePath = string.Empty;

        public UploadFileOnRemote(string browser,string remoteIp, string filePath)
        {
            _RemoteIp = remoteIp;
            _browserName = browser;
            _FilePath = filePath;
            if (_browserName.ToLower().Equals(Common.KryptonConstants.BROWSER_FIREFOX))
            {             
                _AutoitExeName = "FileUploadFF.exe";
            }
            else if (_browserName.ToLower().Equals(Common.KryptonConstants.BROWSER_CHROME))
            {
               _AutoitExeName = "FileUploadChrome.exe";
            }
            else
            {               
                _AutoitExeName = "FileUploadIE.exe";
            }

        }

        public void UploadFileWithAutoIt()
        {
            try
            {
                ServiceAgent service = new ServiceAgent(_RemoteIp); //Node ip          
                service.StartExecutableByName(_AutoitExeName, _FilePath);
                Console.WriteLine(string.Format("Uploading file  {0}  on  {1}  machine..!!", _FilePath, _RemoteIp));
            }
            catch(Exception ex)
            {
                throw new Exception( string.Format("Coud not Upload File on {0} machine. Actual Error:{1}",_RemoteIp,ex.Message));
            }

            


        }


    }
}

