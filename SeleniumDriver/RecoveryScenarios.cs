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
using System.Linq;
using System.Text;
using System.Data;
using Common;

namespace Driver
{
    public class RecoveryScenarios
    {
        private static DataSet recoverData = new DataSet();
        private static DataSet datasetRecoverBrowser = new DataSet();
        private static DataSet datasetOR = new DataSet();
        private static int[] noOfTimes;
        private static Dictionary<string, string> objDataRow = new Dictionary<string, string>();
        private static ITestObject objTestObject = null;
        private static string cachedAttribute = String.Empty;
        private static string cachedAttributeType = String.Empty;
        //Constructor for collecting all data.
        public RecoveryScenarios(DataSet recoverDataset, DataSet recoverData, DataSet ORData, ITestObject objTest)
        {
            RecoveryScenarios.recoverData = recoverDataset;
            RecoveryScenarios.datasetRecoverBrowser = recoverData;
            RecoveryScenarios.datasetOR = ORData;
            objTestObject = objTest;
            try
            {
                if(recoverData.Tables.Count>0)
                RecoveryScenarios.noOfTimes = new int[recoverData.Tables[0].Rows.Count];
            }
            catch (System.IndexOutOfRangeException)
            {

            }
        }


        public static void cacheAttribute(string attributeType, string attribute)
        {
            cachedAttribute = attribute;
            cachedAttributeType = attributeType;
        }
        //Recover all win32 alerts ,ie. Accept or Dismiss.
        public void recoverFromPopUps()
        {
            bool handled = false;
            string actualPopupText = string.Empty;
            try
            {
                actualPopupText = Browser.driver.SwitchTo().Alert().Text;
                for (int i = 0; i < recoverData.Tables[0].Rows.Count; i++)
                {
                    string expectedPopupText = recoverData.Tables[0].Rows[i]["PopUpText"].ToString().Trim();
                    if (actualPopupText.Contains(expectedPopupText))
                    {
                        switch (recoverData.Tables[0].Rows[i]["Action"].ToString().ToLower().Trim())
                        {
                            case "accept":
                                Browser.driver.SwitchTo().Alert().Accept();
                                break;
                            case "dismiss":
                                Browser.driver.SwitchTo().Alert().Dismiss();
                                break;
                            default:
                                Exception e = new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0017"));
                                throw e;
                        }
                        handled = true;
                        recoverFromPopUps();
                    }
                }
                if (!handled)
                {
                    Exception popupEx = new Exception(Common.Utility.GetCommonMsgVariable("KRYPTONERRCODE0018").Replace("{MSG}", actualPopupText));
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
        public static void recoverFromBrowsers(DataSet recoverDataset = null, DataSet brRecoverData = null, DataSet orDataSet = null, TestObject tObject = null)
        {
            //This indicates if an actual recovery was performed
            if (recoverDataset != null && brRecoverData != null && orDataSet != null && tObject != null)
            {
                recoverData = recoverDataset;
                datasetRecoverBrowser = brRecoverData;
                datasetOR = orDataSet;
                objTestObject = tObject;
                RecoveryScenarios.noOfTimes = new int[brRecoverData.Tables[0].Rows.Count];
            }
           
            try
            {
 				 
                for (int i = 0; i < datasetRecoverBrowser.Tables[0].Rows.Count; i++)
                {
                    Common.Property.isRecoveryRunning = true;
                    if (noOfTimes[i].Equals(Common.Property.RECOVERY_COUNT))
                    {
                        break;
                    }
                    string keyword = datasetRecoverBrowser.Tables[0].Rows[i]["Recovery_Keyword"].ToString().Trim();
                    string action = datasetRecoverBrowser.Tables[0].Rows[i]["Action"].ToString().Trim();
                    string recoverDetails = datasetRecoverBrowser.Tables[0].Rows[i]["Recovery_Details"].ToString().Trim();
                    string data = string.Empty;
                    try
                    {
                        data = datasetRecoverBrowser.Tables[0].Rows[i]["Data"].ToString().Trim();
                    }
                    catch (Exception) { }

                    switch (keyword.ToLower())
                    {
                        case "object_presence":
                            string[] recover_Contents = recoverDetails.Split('|');
                            string parent = recover_Contents[0].Trim();
                            string testObject = recover_Contents[1].Trim();
                            objDataRow = Common.Utility.GetTestOrData(parent, testObject, datasetOR);
                            if (objDataRow.Count.Equals(0))
                            {
                                continue;
                            }
                            try
                            {
                                TestObject.attribute = objDataRow[KryptonConstants.WHAT];
                                TestObject.attributeType = objDataRow[KryptonConstants.HOW];
                                //Will Add more actions later, for now click is in demand.
                                switch (action.ToLower())
                                {
                                    case "click":
                                        objTestObject.Click();
                                        Console.WriteLine("Recovery Active: Click on " + parent + " | " + testObject);
                                        break;
                                    case "fireevent":
                                        objTestObject.FireEvent(data);
                                        Console.WriteLine("Recovery Active: FireEvent: " + data + " on " + parent + " | " + testObject);
                                        break;
                                    default:
                                        break;
                                }
                                //KRYPTON0419.
                                Browser.driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(3.0));
                                noOfTimes[i]++;
                            }
                            catch (Exception)
                            {
                                noOfTimes[i] = 0;
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
                TestObject.attribute = cachedAttribute;
                TestObject.attributeType = cachedAttributeType;
                Common.Property.isRecoveryRunning = false;
            }
            catch (Exception e)
            {
                Common.Property.isRecoveryRunning = false;
            }
        }
        #endregion
    }
}
