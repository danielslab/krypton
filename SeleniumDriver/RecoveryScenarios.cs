/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Krypton.TestEngine.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Main Engine which interact with all other components and drive them.
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using Common;
using Driver.Browsers;

namespace Driver
{
    public class RecoveryScenarios
    {
        private static DataSet _recoverData = new DataSet();
        private static DataSet _datasetRecoverBrowser = new DataSet();
        private static DataSet _datasetOr = new DataSet();
        private static int[] _noOfTimes;
        private static Dictionary<string, string> _objDataRow = new Dictionary<string, string>();
        private static ITestObject _objTestObject = null;
        private static string _cachedAttribute = String.Empty;
        private static string _cachedAttributeType = String.Empty;
        //Constructor for collecting all data.
        public RecoveryScenarios(DataSet recoverDataset, DataSet recoverData, DataSet orData, ITestObject objTest)
        {
            _recoverData = recoverDataset;
            _datasetRecoverBrowser = recoverData;
            _datasetOr = orData;
            _objTestObject = objTest;
            try
            {
                if(recoverData.Tables.Count>0)
                _noOfTimes = new int[recoverData.Tables[0].Rows.Count];
            }
            catch (IndexOutOfRangeException)
            {

            }
        }


        public static void CacheAttribute(string attributeType, string attribute)
        {
            _cachedAttribute = attribute;
            _cachedAttributeType = attributeType;
        }
        //Recover all win32 alerts ,ie. Accept or Dismiss.
        public void RecoverFromPopUps()
        {
            bool handled = false;
            try
            {
                var actualPopupText = Browser.Driver.SwitchTo().Alert().Text;
                for (int i = 0; i < _recoverData.Tables[0].Rows.Count; i++)
                {
                    string expectedPopupText = _recoverData.Tables[0].Rows[i]["PopUpText"].ToString().Trim();
                    if (actualPopupText.Contains(expectedPopupText))
                    {
                        switch (_recoverData.Tables[0].Rows[i]["Action"].ToString().ToLower().Trim())
                        {
                            case "accept":
                                Browser.Driver.SwitchTo().Alert().Accept();
                                break;
                            case "dismiss":
                                Browser.Driver.SwitchTo().Alert().Dismiss();
                                break;
                            default:
                                Exception e = new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0017"));
                                throw e;
                        }
                        handled = true;
                        RecoverFromPopUps();
                    }
                }
                if (!handled)
                {
                    Exception popupEx = new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0018").Replace("{MSG}", actualPopupText));
                    throw popupEx;

                }
            }
            catch (Exception e)
            {
                throw e;
            }


        }

        #region Recover From Browser Popup functionality goes here.
        /// <summary>
        /// RecoverFromBrowsers handles browser level recovery.
        /// </summary>
        public static void RecoverFromBrowsers(DataSet recoverDataset = null, DataSet brRecoverData = null, DataSet orDataSet = null, TestObject tObject = null)
        {
            //This indicates if an actual recovery was performed
            if (recoverDataset != null && brRecoverData != null && orDataSet != null && tObject != null)
            {
                _recoverData = recoverDataset;
                _datasetRecoverBrowser = brRecoverData;
                _datasetOr = orDataSet;
                _objTestObject = tObject;
                _noOfTimes = new int[brRecoverData.Tables[0].Rows.Count];
            }
           
            try
            {
 				 
                for (int i = 0; i < _datasetRecoverBrowser.Tables[0].Rows.Count; i++)
                {
                    Property.IsRecoveryRunning = true;
                    if (_noOfTimes[i].Equals(Property.RecoveryCount))
                    {
                        break;
                    }
                    string keyword = _datasetRecoverBrowser.Tables[0].Rows[i]["Recovery_Keyword"].ToString().Trim();
                    string action = _datasetRecoverBrowser.Tables[0].Rows[i]["Action"].ToString().Trim();
                    string recoverDetails = _datasetRecoverBrowser.Tables[0].Rows[i]["Recovery_Details"].ToString().Trim();
                    string data = string.Empty;
                    try
                    {
                        data = _datasetRecoverBrowser.Tables[0].Rows[i]["Data"].ToString().Trim();
                    }
                    catch (Exception)
                    {
                        // ignored
                    }

                    switch (keyword.ToLower())
                    {
                        case "object_presence":
                            string[] recoverContents = recoverDetails.Split('|');
                            string parent = recoverContents[0].Trim();
                            string testObject = recoverContents[1].Trim();
                            _objDataRow = Utility.GetTestOrData(parent, testObject, _datasetOr);
                            if (_objDataRow.Count.Equals(0))
                            {
                                continue;
                            }
                            try
                            {
                                TestObject.Attribute = _objDataRow[KryptonConstants.WHAT];
                                TestObject.AttributeType = _objDataRow[KryptonConstants.HOW];
                                //Will Add more actions later, for now click is in demand.
                                switch (action.ToLower())
                                {
                                    case "click":
                                        _objTestObject.Click();
                                        Console.WriteLine("Recovery Active: Click on " + parent + " | " + testObject);
                                        break;
                                    case "fireevent":
                                        _objTestObject.FireEvent(data);
                                        Console.WriteLine("Recovery Active: FireEvent: " + data + " on " + parent + " | " + testObject);
                                        break;
                                    default:
                                        break;
                                }
                                //KRYPTON0419.
                                Browser.Driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(3.0));
                                _noOfTimes[i]++;
                            }
                            catch (Exception)
                            {
                                _noOfTimes[i] = 0;
                                continue;
                            }
                            break;
                        case "browser_title":
                            break;
                        case "browser_url":
                            break;
                    }
                }
                //Resetting attribute and attributeType.
                TestObject.Attribute = _cachedAttribute;
                TestObject.AttributeType = _cachedAttributeType;
                Property.IsRecoveryRunning = false;
            }
            catch (Exception e)
            {
                Property.IsRecoveryRunning = false;
            }
        }
        #endregion
    }
}
