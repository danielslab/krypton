using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Common;
using System.IO;
using System.Data;

namespace Krypton
{
    class KryptonUtility : Variables
    {

        /// <summary>
        /// Return testCaseIds those are present in the Excel
        /// after comparing the suite file or test case id from the Parameter.ini
        /// </summary>
        /// <param></param> 
        /// <returns string = testcaseId>Matched Test Case id in the File</returns>
        internal static string getTestCaseIDs()
        {
            if (!testSuite.Trim().IsNullOrWhiteSpace() && Common.Utility.GetParameter("TestCaseId").Trim().IsNullOrWhiteSpace() || true)
            {
                //If a test suite was specified, check if any keyword was also passed to filter execution
                IEnumerable<string> keywordFilter = Common.Utility.GetParameter("keyword").Trim().Replace(" ", string.Empty).Split(',');
                //Test Id filter. Applicable when user passes one or more test case id
                IEnumerable<string> testFilter = Common.Utility.GetParameter("TestCaseId").Trim().ToLower().Split(',');
                //get managertype from property.cs
                Krypton.Manager.SetTestFilesLocation(Common.Property.ManagerType);
                Krypton.Manager.isTestSuite = true;
                Krypton.Manager objTestManagerSuite = new Krypton.Manager(testSuite);
                DataSet testSuiteData = new DataSet();
                testSuiteData = objTestManagerSuite.GetTestCaseXml(Property.TestQuery, false, false);

                for (int testCaseCnt = 0; testCaseCnt < testSuiteData.Tables[0].Rows.Count; testCaseCnt++)
                {
                    if (!testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString().Trim().IsNullOrWhiteSpace())
                    {

                        string testIdFromXml = testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString().Trim().ToLower();

                        //Collect keyword information specified with a test case
                        string testKeyword = testSuiteData.Tables[0].Rows[testCaseCnt]["keyword"].ToString() + ",";
                        IEnumerable<string> testKeywords = testKeyword.Replace(" ", string.Empty).Split(',');
                        bool keywordsMatch = false;

                        //include test case for execution if even one matching filter is found
                        if (!testKeywords.Intersect(keywordFilter).Count().Equals(0))
                        {
                            keywordsMatch = true;
                        }
                        bool skipTestCase = false;
                        //Check if only a specified number of test cases should be executed
                        if (testFilter.Contains(testIdFromXml) || testFilter.Contains(string.Empty))
                        {
                            skipTestCase = false;
                            m_lstInputTestCaseIDs.Add(testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString().Trim());
                        }
                        else
                        {
                            skipTestCase = true;
                        }
                        //Check modifier information. If a test case contains {skip} on first record, 
                        //that test case should not be executed while running in batch mode
                        string options = testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.OPTIONS].ToString();
                        if (options.ToLower().Contains("{skip}"))
                        {
                            skipTestCase = true;
                        }
                        //Check if test case also includes initialRelease parameter
                        if (options.ToLower().Contains("{initialrelease"))
                        {
                            int initialRelease = Utility.GetNumberFromText(options.ToLower().Split(
                                                 new[] { "{initialrelease" }, StringSplitOptions.None).Last().Split('}').First());
                            int currentRelease = Utility.GetNumberFromText(Utility.GetParameter("currentRelease"));

                            if (initialRelease > currentRelease)
                            {
                                skipTestCase = true;
                            }
                        }
                        //Check if passed on filter matches with keywords in test case file
                        if (keywordsMatch && !skipTestCase)
                        {
                            if (testCaseId.IsNullOrWhiteSpace())
                                testCaseId = testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString();
                            else
                                testCaseId = testCaseId + "," +
                                             testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString();
                        }
                    }
                }
            }
            return testCaseId;
        }



        /// <summary>
        /// Set the all file location into the vairables in the property.cs 
        /// To identify the location of all the project files.
        /// </summary>
        /// <param></param> 
        /// <returns></returns>
        internal static void SetProjectFilesPaths()
        {
            //Get Data File path
            if (Path.IsPathRooted(Common.Utility.GetParameter("TestDataLocation")))
                Common.Property.TestDataFilepath = Common.Utility.GetParameter("TestDataLocation");
            else
                Common.Property.TestDataFilepath = Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("TestDataLocation"));

            //Get DBTestData File path
            if (Path.IsPathRooted(Common.Utility.GetParameter("DBTestDataLocation")))
                Common.Property.DBTestDataFilepath = Common.Utility.GetParameter("DBTestDataLocation");
            else
                Common.Property.DBTestDataFilepath = Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("DBTestDataLocation"));

            //get reusable path
            if (Path.IsPathRooted(Common.Utility.GetParameter("DBTestDataLocation")))
                Common.Property.ReusableLocation = Common.Utility.GetParameter("ReusableLocation");
            else
                Common.Property.ReusableLocation = Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("ReusableLocation"));

            //Get OR File path
            if (Path.IsPathRooted(Common.Utility.GetParameter("ORLocation")))
                Common.Property.ObjectRepositoryFilepath = Common.Utility.GetParameter("ORLocation");
            else
                Common.Property.ObjectRepositoryFilepath = Path.GetFullPath(Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("ORLocation")));

            #region RecoverFromPopuppath.
            if (Path.IsPathRooted(Common.Utility.GetParameter("RecoverFromPopupLocation")))
                Common.Property.RecoverFromPopupFilepath = Common.Utility.GetParameter("RecoverFromPopupLocation");
            else
                Common.Property.RecoverFromPopupFilepath = Path.GetFullPath(string.Concat(Common.Property.ApplicationPath, Common.Utility.GetParameter("RecoverFromPopupLocation")));

            //Get RecoverFromPopup File 
            Property.RecoverFromPopupFilename = Utility.GetParameter("RecoverFromPopupFileName") + ".xlsx";
            Common.Property.MaxTimeoutForPageLoad = int.Parse(Common.Utility.GetParameter("MaxTimeoutForPageLoad"));
            Common.Property.MinTimeoutForPageLoad = int.Parse(Common.Utility.GetParameter("MinTimeoutForPageLoad"));
            if (Path.IsPathRooted(Common.Utility.GetParameter("RecoverFromBrowserLocation")))
                Common.Property.RecoverFromBrowserFilePath = Common.Utility.GetParameter("RecoverFromBrowserLocation");
            else
                Common.Property.RecoverFromBrowserFilePath = Path.GetFullPath(string.Concat(Common.Property.ApplicationPath, Common.Utility.GetParameter("RecoverFromBrowserLocation")));

            Property.RecoverFromBrowserFilename = Utility.GetParameter("RecoverFromBrowserFileName") + ".xlsx";
            #endregion
            //Get Company logo path
            if (Path.IsPathRooted(Common.Utility.GetParameter("CompanyLogo")))
                Common.Property.CompanyLogo = Common.Utility.GetParameter("CompanyLogo");
            else
                Common.Property.CompanyLogo = Path.GetFullPath(Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("CompanyLogo")));
            if (!File.Exists(Common.Property.CompanyLogo))
                Common.Property.CompanyLogo = string.Empty;

            if (Path.IsPathRooted(Common.Utility.GetParameter("parallelRecovery")))
                Common.Property.ParallelRecoveryFilePath = Common.Utility.GetParameter("parallelRecovery");
            else
                Common.Property.ParallelRecoveryFilePath = Path.GetFullPath(Path.Combine(Common.Property.ApplicationPath, Common.Utility.GetParameter("parallelRecovery")));
            //Database Connection String initialization
            Common.Property.DbConnectionString = Common.Utility.GetParameter("DBConnectionString");

            //Driver Error capture as image or html
            Common.Property.ErrorCaptureAs = Common.Utility.GetParameter("ErrorCaptureAs");

            //extract parallel recoverysheet name from parameter.ini
            if (Utility.GetParameter("ParallelRecoverySheetName").Trim().Length > 0)
            {
                Property.ParallelRecoverySheetName = Utility.GetParameter("ParallelRecoverySheetName") + ".xlsx";
            }
            else
            {
                Property.ParallelRecoverySheetName = "popuprecovery" + Property.ExcelSheetExtension;
            }

        }
    }
}
