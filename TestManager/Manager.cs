/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Krypton.Manager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Main file to retrive data from excel file or execute method to fetch excel file
** from xStudio/QC/TFS etc. It also store log result to a predefined location
*****************************************************************************/

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using Common;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using ExcelLib;

namespace Krypton
{
    public class Manager
    {
        private static string strTestCaseId = string.Empty;
        private static string strManagerType = string.Empty;

        private static string testCaseFilePath = string.Empty;
        private static string testDataFilePath = string.Empty;
        private static string dbTestDataFilePath = string.Empty;
        private static string objectRepositoryFilePath = string.Empty;
        private static string recoverFromPopupFilePath = string.Empty;
        private static string recoverFromBrowserFilePath = string.Empty;
        private static DirectoryInfo strResultsSource;
        private static DirectoryInfo strResultsDestination;

        private static string resultFilesList = string.Empty;

        public static bool isTestSuite = false;

        private static string newTestCaseId = string.Empty;



        public Manager(string testCaseId, string testCasefilename = null)
        {
             char strSeprator = char.Parse(Property.TestCaseIDSeperator);
                strTestCaseId = testCaseId;
            SetTestFilesLocation(strManagerType, testCasefilename);

            newTestCaseId = string.Empty;
            if (strTestCaseId.Contains("TC_"))
            {
            newTestCaseId = strTestCaseId.Remove(0, 3);
            }
            else
            {
                newTestCaseId = strTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter];
            }
            //Store test case id to current test case variable
            Common.Property.CurrentTestCase = strTestCaseId;
        }

        /// <summary>
        /// This method is used to initiate execution in test manager, create test run and results and set required parameters
        /// </summary>
        /// <returns></returns>
        public static void InitTestCaseExecution()
        {
            //Send instructions to individual test managers for initialization of execution
            //Current implementation only for Visual Studio test manager, should be extended to others also
            try
            {
                switch (Common.Property.ManagerType.ToLower())
                {
                    case "mstestmanager":
                        TestManager MSTestManager = new TestManager();
                        MSTestManager.InitExecution();
                        break;
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        ///  This method will set the path of the the test case, test data & object repository files
        /// according to the test manager (FileSystem, QC, XStudio etc.) used by the Krypton.
        /// </summary>
        public static void SetTestFilesLocation(string managerType, string testCasefilename = null)
        {
            strManagerType = managerType;
            char strSeprator = char.Parse(Property.TestCaseIDSeperator);
            switch (strManagerType.ToLower())
            {
                //For file system, it will get the test files paths from property.cs
                case "filesystem":
                    if (isTestSuite)
                    {
                        testCaseFilePath =Path.GetFullPath( Property.TestCaseFilepath + "/" + strTestCaseId + Property.ExcelSheetExtension);
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(testCasefilename))
                            testCaseFilePath =Path.GetFullPath(Property.TestCaseFilepath + "/" + strTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension);
                        else
                            testCaseFilePath = Path.GetFullPath(Property.TestCaseFilepath + "/" + testCasefilename);
                    }
                    if (string.IsNullOrWhiteSpace(testCasefilename))
                        testDataFilePath = Path.GetFullPath(Property.TestDataFilepath + "/" + "TD_" + newTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension);//New
                    else
                    {
                        testDataFilePath = Path.GetFullPath(Property.TestDataFilepath + "/" + "TD_" + testCasefilename);
                    }

                        dbTestDataFilePath = Path.GetFullPath(Property.DBTestDataFilepath + "/" + "DB_" + newTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension);
                    objectRepositoryFilePath = Path.GetFullPath(Property.ObjectRepositoryFilepath + "/" + Property.ObjectRepositoryFilename);
                    recoverFromPopupFilePath = Path.GetFullPath( Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename);
                    recoverFromBrowserFilePath =Path.GetFullPath(Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename);

                    break;
                //for QC, it will get the test files paths from the TMQC module.
                case "qc":
                    TestManager objExtManager = new TestManager();
                    //Call DownloadAttachment from TMQC TestManager class to get the path of test files.
                    if (isTestSuite)
                    {
                        testCaseFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], strTestCaseId + Property.ExcelSheetExtension);

                    }
                    else
                    {

                        if (string.IsNullOrWhiteSpace(testCasefilename))
                            testCaseFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], Regex.Split(strTestCaseId, Property.TestCaseIDSeperator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension);
                        else
                            testCaseFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], testCasefilename);

                    }

                    if (string.IsNullOrWhiteSpace(testCasefilename))
                        testDataFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], Regex.Split(strTestCaseId, Property.TestCaseIDSeperator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension);
                    else
                    {
                        testDataFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], testCasefilename);
                    }

                    dbTestDataFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], Regex.Split(strTestCaseId, Property.TestCaseIDSeperator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension);
                    objectRepositoryFilePath = objExtManager.DownloadAttachment(Common.Property.parameterdic["qcfolder"], Property.ObjectRepositoryFilename);
                    recoverFromPopupFilePath = Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename;
                    recoverFromBrowserFilePath = Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename;

                    break;
                //for QC, it will get the test files paths from the TMXStudio module.
                case "xstudio":
                    if (isTestSuite)
                    {
                        testCaseFilePath = Property.TestCaseFilepath + "/" + strTestCaseId + Property.ExcelSheetExtension;
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(testCasefilename))
                            testCaseFilePath = Property.TestCaseFilepath + "/" + strTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension;
                        else
                            testCaseFilePath = Property.TestCaseFilepath + "/" + testCasefilename;
                    }
                    if (string.IsNullOrWhiteSpace(testCasefilename))
                        testDataFilePath = Property.TestDataFilepath + "/" + strTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension;
                    else
                    {
                        testDataFilePath = Property.TestDataFilepath + "/" + testCasefilename;
                    }

                    dbTestDataFilePath = Property.TestDataFilepath + "/" + strTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension;
                    objectRepositoryFilePath = Property.ObjectRepositoryFilepath + "/" + Property.ObjectRepositoryFilename;
                    recoverFromPopupFilePath = Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename;
                    recoverFromBrowserFilePath = Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename;

                    break;
                // For Microsoft Test Manager
                case "mstestmanager":
                    // Not Functional yet.
                    if (isTestSuite)
                    {
                        testCaseFilePath = Property.TestCaseFilepath + "/" + strTestCaseId + Property.ExcelSheetExtension;
                    }
                    else
                    {

                        if (string.IsNullOrWhiteSpace(testCasefilename))
                            testCaseFilePath = Property.TestCaseFilepath + "/" + strTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension;
                        else
                            testCaseFilePath = Property.TestCaseFilepath + "/" + testCasefilename;
                    }
                    if (string.IsNullOrWhiteSpace(testCasefilename))
                      testDataFilePath = Property.TestDataFilepath + "/" + "TD_" + newTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension;
                    else
                    {
                        testDataFilePath = Property.TestDataFilepath + "/" + "TD_" + testCasefilename;
                    }

                    dbTestDataFilePath = Property.DBTestDataFilepath + "/" + "DB_" + newTestCaseId.Split(strSeprator)[Property.TestCaseIDParameter] + Property.ExcelSheetExtension; 
                    objectRepositoryFilePath = Property.ObjectRepositoryFilepath + "/" + Property.ObjectRepositoryFilename;
                    recoverFromPopupFilePath = Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename;
                    recoverFromBrowserFilePath = Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename;

                    break;
            }

        }

        
        /// <summary>
        /// This method will fetch test steps from Test Case Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Case Steps.</returns>
        public DataSet GetTestCaseXml(string testFlowQuery,bool IsGlobalReusabale,bool IsSpecificSheet)
        {
            string filePathPath = string.Empty;
            if (IsGlobalReusabale)
            {
                string globalActionFileName = string.Empty;
                if (!Path.GetExtension(Utility.GetParameter("reusabledefaultfile").ToLower().Trim()).Equals(Property.ExcelSheetExtension.ToLower()))
                    globalActionFileName = Utility.GetParameter("reusabledefaultfile") + Property.ExcelSheetExtension;
                else
                    globalActionFileName = Utility.GetParameter("reusabledefaultfile").Trim();

                filePathPath = Path.Combine(Property.ReusableLocation, globalActionFileName);
            
            }
            else if(IsSpecificSheet)            
            {
                filePathPath = Path.Combine(Property.ReusableLocation, testFlowQuery.Trim() + Property.ExcelSheetExtension);
            }
            else
                filePathPath = testCaseFilePath;

            try
            {
                //call GetRequiredRows method to read test steps inside test case sheet.
                DataSet strTestCaseFlow = GetDataSet(filePathPath, "testFlow");
                validateTestCase(strTestCaseFlow);
                return strTestCaseFlow;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        /// <summary>
        /// This method will fetch test data from Test Case Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Data.</returns>
        public DataSet GetTestDataXml()
        {
            string fileName = string.Empty;
            if (Common.Property.TestDataSheet.ToLower().Equals("test_data"))
            {
                fileName = testDataFilePath;
            }
            else
            {
                fileName = Path.Combine(Path.GetDirectoryName(testDataFilePath), Property.TestDataSheet + Property.ExcelSheetExtension);
            }       
            try
            {
                DataSet strTestDataFlow = GetDataSet(fileName, "TestData"); ;
                validateTestData(strTestDataFlow);
                return strTestDataFlow;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        /// <summary>
        /// Get Popup sheet and return as a dataset.
        /// </summary>
        /// <returns></returns>
        public DataSet GetRecoverFromPopupXml()
        {
            try
            {
                DataSet strRecoverFromPopupFlow = GetDataSet(recoverFromPopupFilePath, "recoveryDataSet");        
                validateRecoveryFromPopups(strRecoverFromPopupFlow);
                return strRecoverFromPopupFlow;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        /// <summary>
        ///  Get "IE_Popups" workSheet and return as a dataset.
        /// </summary>
        /// <returns></returns>
       

        /// <summary>
        /// Get Popup sheet and return as a dataset.
        /// </summary>
        /// <ret        urns></returns>
        public DataSet GetRecoverFromBrowserXml()
        {
            try
            {
                DataSet strRecoverFromBrowserFlow = GetDataSet(recoverFromBrowserFilePath, "recoveryDataSet");  
                validateRecoveryFromBrowser(strRecoverFromBrowserFlow);
                return strRecoverFromBrowserFlow;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }
        /// <summary>
        ///This method will fetch DB test data from DBTestData Excel Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Data.</returns>
        public DataSet GetDBTestDataXml()
        {
            try
            {
                DataSet strTestDataFlow = GetDataSet(dbTestDataFilePath, "TestDbDataflow"); 
                validateDBTestData(strTestDataFlow);
                return strTestDataFlow;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        /// <summary>
        ///  This method will fetch objects from Object Repository Excel Sheet into a dataset
        /// and will return the dataset as XML string.
        /// It will get the object repository sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Object Definition.</returns>
        public static DataSet GetObjectDefinitionXml() //static removed
        {
            try
            {
                DataSet strObjectDefinition = GetDataSet(objectRepositoryFilePath, "ORDataSet"); 
                validateOR(strObjectDefinition);
                return strObjectDefinition;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        /// <summary>
        ///  This method will copy report files from a temporary location used by Reporting module
        /// to a given location that can be used by the QC or XStudio or similar test management modules. It also
        /// generates a text file containing all result files' path for XStudio.
        /// Source and destination paths will be fetched from the Property.cs file.
        /// </summary>
        /// <returns>True on success else False</returns>
        public static bool UploadTestExecutionResults(List<string> InputTestIds= null)
        {
            bool res = false;
            try
            {

                strResultsSource = new DirectoryInfo(Property.ResultsSourcePath).Parent;
                strResultsDestination = new DirectoryInfo(Property.ResultsDestinationPath);

                // Copy all result files from source directory to destination directory.
                res = CopyAll(strResultsSource, strResultsDestination);

                // Create a temporary file to make sure that the copy process has been completed.
                bool resFinished = true;
                StreamWriter sw = new StreamWriter(strResultsDestination.Parent.FullName + "/ResultCopyCompleted.txt");
                sw.WriteLine("Kay: Copy Process Completed!");
                sw.Close();

                DirectoryInfo drParent = new DirectoryInfo(Property.ResultsSourcePath).Parent.Parent;

                //to check if user wants to keep report history
                string fileExt = string.Empty;
                if (strResultsDestination.Name.LastIndexOf('-') >= 0 && Common.Utility.GetParameter("KeepReportHistory").Equals("true", StringComparison.OrdinalIgnoreCase) == true)
                    fileExt = strResultsDestination.Name.Substring(strResultsDestination.Name.LastIndexOf('-'),
                                                                   strResultsDestination.Name.Length - strResultsDestination.Name.LastIndexOf('-'));

                string RCProcessId = string.Empty;
                if (!string.IsNullOrWhiteSpace(Common.Property.RCProcessId))
                    RCProcessId = "-" + Common.Property.RCProcessId; //Adding process id for Remote execution

                foreach (FileInfo fi in drParent.GetFiles())
                {
                    if (fi.Name.IndexOf("HtmlReport", StringComparison.OrdinalIgnoreCase) >= 0
                        || fi.Name.IndexOf("logo", StringComparison.OrdinalIgnoreCase) >= 0
                        || fi.Name.IndexOf(Common.Property.ReportZipFileName, StringComparison.OrdinalIgnoreCase) >= 0
                        || fi.Name.IndexOf("." + Property.ScriptLanguage, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string fileName = string.Empty;
                        if (fi.Name.IndexOf("mail.html", StringComparison.OrdinalIgnoreCase) >= 0) // no need to copy this file to krypton result folder while sending mail through krypton
                            continue;
                        if (fi.Name.IndexOf("logo", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            fileName = fi.Name;
                        }
                        else if (fi.Name.IndexOf("HtmlReport", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string htmlFileExt = string.Empty;

                            if (!string.IsNullOrWhiteSpace(Common.Utility.GetParameter("TestSuite")) && string.IsNullOrWhiteSpace(Common.Utility.GetParameter("TestCaseId")))
                            {
                                htmlFileExt = Common.Utility.GetParameter("TestSuite");
                            }
                            else
                            {
                                string[] testCaseIds = Common.Utility.GetParameter("TestCaseId").Split(',');
                                if (testCaseIds.Length > 1) htmlFileExt = testCaseIds[0] + "...multiple";
                                else htmlFileExt = Common.Utility.GetParameter("TestCaseId");
                            }

                            htmlFileExt = htmlFileExt + RCProcessId;

                            if (fi.Name.IndexOf("HtmlReports", StringComparison.OrdinalIgnoreCase) >= 0)
                            fileName = fi.Name.Substring(0, fi.Name.LastIndexOf('.')) +
                                "-" + htmlFileExt + fileExt+"s"+ fi.Name.Substring(fi.Name.LastIndexOf('.'), fi.Name.Length - fi.Name.LastIndexOf('.'));
                            else
                                fileName = fi.Name.Substring(0, fi.Name.LastIndexOf('.')) +
                                "-" + htmlFileExt + fileExt + fi.Name.Substring(fi.Name.LastIndexOf('.'), fi.Name.Length - fi.Name.LastIndexOf('.'));
                        }
                        else
                        {
                            fileName =
                              fi.Name.Substring(0, fi.Name.LastIndexOf('.')) +
                              fileExt +
                              fi.Name.Substring(fi.Name.LastIndexOf('.'), fi.Name.Length - fi.Name.LastIndexOf('.'));
                        }


                        fi.CopyTo(strResultsDestination.Parent.FullName + "/" + fileName, true);

                        //if user wants to keep report history, folder name will be different, so need to change links in the html file also
                        if (string.IsNullOrWhiteSpace(fileExt) == false && fi.Name.Contains("HtmlReport"))
                        {
                            string htmlStr = File.ReadAllText(strResultsDestination.Parent.FullName + "/" + fileName,Encoding.UTF8);

                            string testCaseId = string.Empty;

                            //Handle Test Case Id conditions
                            #region check for test suite
                            string testSuite = Common.Utility.GetParameter("TestSuite");
                            if (!string.IsNullOrWhiteSpace(testSuite.Trim()) && string.IsNullOrWhiteSpace(Common.Utility.GetParameter("TestCaseId").Trim()))
                            {
                                Krypton.Manager.SetTestFilesLocation(Common.Property.ManagerType); // get ManagerType from property.cs
                                Krypton.Manager objTestManagerSuite = new Krypton.Manager(testSuite);
                                DataSet testSuiteData = new DataSet();
                                testSuiteData = objTestManagerSuite.GetTestCaseXml(Property.TestQuery,false,false);
                                for (int testCaseCnt = 0; testCaseCnt < testSuiteData.Tables[0].Rows.Count; testCaseCnt++)
                                {
                                    if (!string.IsNullOrWhiteSpace(testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString().Trim()))
                                    {
                                        if (string.IsNullOrWhiteSpace(testCaseId))
                                            testCaseId = testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString();
                                        else
                                            testCaseId = testCaseId + "," +
                                                         testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString();
                                    }
                                }

                            }
                            else
                            {
                                testCaseId = Common.Utility.GetParameter("TestCaseId");//override testcaseid parameter over test suite
                            }

                            #endregion

                            string[] testCases=null;
                            if(InputTestIds!=null)
                                testCases = InputTestIds.ToArray();
                            else
                            testCases = testCaseId.Split(',');

                            string[] testCasesIds = new string[testCases.Length];

                            int k = 0;
                            for (int testCaseCnt = 0; testCaseCnt < testCases.Length; testCaseCnt++)
                            {
                                if (string.Join(";", testCasesIds).Contains(testCases[testCaseCnt]+";") == false)
                                {
                                    testCasesIds[k] = testCases[testCaseCnt];
                                    htmlStr = htmlStr.Replace("href='" + testCases[testCaseCnt]+ "\\",
                                                              "href='" + testCases[testCaseCnt] + RCProcessId + fileExt+ "\\");
                                    k++;
                                }
                            }


                            File.WriteAllText(strResultsDestination.Parent.FullName + "/" + fileName, htmlStr,Encoding.UTF8);

                        }

                    }
                }
                int resCheckLimit = 0;
                do
                {
                    if (File.Exists(strResultsDestination.Parent.FullName + "/ResultCopyCompleted.txt"))
                    {
                        File.Delete(strResultsDestination.Parent.FullName + "/ResultCopyCompleted.txt");
                        resFinished = false;
                    }
                    else
                    {
                        if (resCheckLimit++ < 6)
                            Thread.Sleep(5000);
                        else
                            resFinished = false;
                    }
                }
                while (resFinished);

            }
            catch (Exception exception)
            {
                throw new Common.KryptonException("error", exception.Message);
            }

            string fileUploadPath = string.Empty;
            fileUploadPath = Property.ResultsDestinationPath + "/" + (new DirectoryInfo(Property.ResultsSourcePath).Name);

            // Extra actions are performed for particular test manager.
            switch (strManagerType.ToLower())
            {
                case "qc":
                    // Not Functional yet
                    break;

                case "xstudio":
                    // XStudio requires a trigger that tells about the end of the execution.
                    // Generation of test_completed.txt file will notify XStudio.
                    try
                    {
                        TestManager.UploadTestResults(resultFilesList, fileUploadPath);
                    }
                    catch (Exception exception)
                    {
                        throw exception;
                    }
                    break;

                case "mstestmanager":
                    // Not Functional yet
                    try
                    {
                        TestManager.UploadTestResults(Common.Property.finalXmlPath, Property.FinalExecutionStatus);
                    }
                    catch (Exception exception)
                    {
                        throw exception;
                    }
                    break;

                default:
                    // By default FileSystem is considered as test manager.
                    break;
            }

            return res;
        }


        /// <summary>
        ///  This method copies the content of a given directory into a destination directory.
        /// All the contents into the source folder including folder hierarchies and files are copied to the
        /// destination directory.
        /// It returns true on a successful copy process.
        /// </summary>
        /// <param name="source">Source Directory Path</param>
        /// <param name="target">Destination Directory Path</param>
        /// <returns>True on success else False</returns>
        public static bool CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            try
            {
                //check whether target directory exists and create it if not present.
                if (Directory.Exists(target.FullName) == false)
                {
                    Directory.CreateDirectory(target.FullName);
                }
                resultFilesList = string.Empty;
                //copy all the files into source directory to target directory.
                foreach (FileInfo fi in source.GetFiles())
                {
                    fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);

                    //this string stores the filenames along with full path for XStudio test_completed.txt file.
                    resultFilesList += (fi.Name.Split('.')[1] == "xml") ? "Result_File=" + Path.Combine(target.ToString(), fi.Name) + "\r\n" : "Upload_File=" + Path.Combine(target.ToString(), fi.Name) + "\r\n";
                }

                //find all the directories into source directory and make a recursive call to copy its file to destination folder.
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                    CopyAll(diSourceSubDir, nextTargetSubDir);
                }
            }
            catch (Exception exception)
            {
                throw new Common.KryptonException("error", exception.Message);
            }

            return true;
        }

        /// <summary>
        ///  Validate the column names for test case dataset.
        /// </summary>
        /// <param name="testCaseXml">DataSet of the current test case sheet to be executed</param>
        public static void validateTestCase(DataSet testCaseXml)
        {
            bool colFound = false;
            string[] testCaseColumns = { "keyword",
                                         "test_scenario",
                                         KryptonConstants.TEST_CASE_ID,
                                         "comments",
                                         "parent",
                                         KryptonConstants.TEST_OBJECT,
                                         "step_action",
                                         "data",
                                         "iteration",
                                         "options" };

            // Verify each column name in test case sheet.
            try
            {
                for (int i = 0; i < testCaseColumns.GetLength(0); i++)
                {
                    colFound = false;
                    for (int j = 0; j < testCaseXml.Tables[0].Columns.Count; j++)
                    {
                        if (testCaseColumns[i] == testCaseXml.Tables[0].Columns[j].ColumnName.ToString().ToLower())
                        {
                            colFound = true;
                            break;
                        }
                    }
                    if (!colFound)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0038").Replace("{MSG}", testCaseColumns[i]));
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0031").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for test data dataset.
        /// </summary>
        /// <param name="testDataXml">DataSet of the current test data sheet</param>
        public static void validateTestData(DataSet testDataXml)
        {
            bool colFound = false;
            string[] testDataColumns = { "row_no" };

            // Verify each column name in test data sheet.
            try
            {
                for (int i = 0; i < testDataColumns.GetLength(0); i++)
                {
                    colFound = false;
                    for (int j = 0; j < testDataXml.Tables[0].Columns.Count; j++)
                    {
                        if (testDataColumns[i] == testDataXml.Tables[0].Columns[j].ColumnName.ToString().ToLower())
                        {
                            colFound = true;
                            break;
                        }
                    }
                    if (!colFound)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0038").Replace("{MSG}", testDataColumns[i]));
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0032").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for object repository dataset.
        /// </summary>
        /// <param name="objectRepositoryXml">DataSet of the current object repository sheet</param>
        public static void validateOR(DataSet objectRepositoryXml)
        {
            bool colFound = false;
            string[] objectRepositoryColumns = { "sl_no",
                                                 "parent",
                                                 KryptonConstants.TEST_OBJECT,
                                                 "logical_name",
                                                 "locale",
                                                 KryptonConstants.OBJ_TYPE,
                                                 KryptonConstants.HOW,
                                                 KryptonConstants.WHAT,
                                                 "comments",
                                                 KryptonConstants.MAPPING };

            //Verify each column name in object repository sheet.
            try
            {
                for (int i = 0; i < objectRepositoryColumns.GetLength(0); i++)
                {
                    colFound = false;
                    for (int j = 0; j < objectRepositoryXml.Tables[0].Columns.Count; j++)
                    {
                        if (objectRepositoryColumns[i] == objectRepositoryXml.Tables[0].Columns[j].ColumnName.ToString().ToLower())
                        {
                            colFound = true;
                            break;
                        }
                    }
                    if (!colFound)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0038").Replace("{MSG}", objectRepositoryColumns[i]));
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0033").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for db test data dataset.
        /// </summary>
        /// <param name="dbTestDataXml">DataSet of the current db test data sheet</param>
        public static void validateDBTestData(DataSet dbTestDataXml)
        {
            bool colFound = false;
            string[] dbTestDataColumns = { "sno",
                                           "sqlquery" };

            // Verify each column name in test data sheet.
            try
            {
                for (int i = 0; i < dbTestDataColumns.GetLength(0); i++)
                {
                    colFound = false;
                    for (int j = 0; j < dbTestDataXml.Tables[0].Columns.Count; j++)
                    {
                        if (dbTestDataColumns[i] == dbTestDataXml.Tables[0].Columns[j].ColumnName.ToString().ToLower())
                        {
                            colFound = true;
                            break;
                        }
                    }
                    if (!colFound)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0038").Replace("{MSG}", dbTestDataColumns[i]));
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0034").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for recovery from popup dataset.
        /// </summary>
        /// <param name="testCaseXml">DataSet of the recovery from popup sheet</param>
        public static void validateRecoveryFromPopups(DataSet recoveryFromPopupsXml)
        {
            bool colFound = false;
            string[] recoveryFromPopupsColumns = { "popuptext",
                                                   "action" };

            // Verify each column name in PopUp Recovery sheet.
            try
            {
                for (int i = 0; i < recoveryFromPopupsColumns.GetLength(0); i++)
                {
                    colFound = false;
                    for (int j = 0; j < recoveryFromPopupsXml.Tables[0].Columns.Count; j++)
                    {
                        if (recoveryFromPopupsColumns[i] == recoveryFromPopupsXml.Tables[0].Columns[j].ColumnName.ToString().ToLower())
                        {
                            colFound = true;
                            break;
                        }
                    }
                    if (!colFound)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0038").Replace("{MSG}", recoveryFromPopupsColumns[i]));
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0035").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for recovery from browser dataset.
        /// </summary>
        /// <param name="testCaseXml">DataSet of the recovery from browser sheet</param>
        public static void validateRecoveryFromBrowser(DataSet recoveryFromBrowserXml)
        {
            bool colFound = false;
            string[] recoveryFromBrowserColumns = { "recovery_keyword",
                                                    "recovery_details",
                                                    "action" };

            // Verify each column name in Browser Recovery sheet.
            try
            {
                for (int i = 0; i < recoveryFromBrowserColumns.GetLength(0); i++)
                {
                    colFound = false;
                    for (int j = 0; j < recoveryFromBrowserXml.Tables[0].Columns.Count; j++)
                    {
                        if (recoveryFromBrowserColumns[i] == recoveryFromBrowserXml.Tables[0].Columns[j].ColumnName.ToString().ToLower())
                        {
                            colFound = true;
                            break;
                        }
                    }
                    if (!colFound)
                        throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0038").Replace("{MSG}", recoveryFromBrowserColumns[i]));
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0036").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  To generate a temp file and return file path
        /// </summary>
        /// <param name="oriFilePath"></param>
        /// <returns></returns>
        private static string GetTempFile(string oriFilePath)
        {
            string tmpFileName = string.Empty;
            tmpFileName = Utility.GetTemporaryFile(Path.GetExtension(oriFilePath));
            if (File.Exists(oriFilePath))
            {
                File.Copy(oriFilePath, tmpFileName);
            }
            else
            {
                throw new Common.KryptonException("Error", Utility.GetCommonMsgVariable("KRYPTONERRCODE0037").Replace("{MSG}", oriFilePath));
            }
            return tmpFileName;
        }

        /// <summary>
        ///  This method read an csv file and returns a dataset.
        /// </summary>
        /// <param filePath="tempfilepath"></param>
        ///  <param dataSetname="tempfilepath"></param>
        /// <returns></returns>
        private static DataSet GetDataSet(string realfilePath, string dataSetname = "dataset")
        {
            DataSet ds = new DataSet(dataSetname);
            string tmpFileName = GetTempFile(realfilePath);
         try
         {
             if (Path.GetExtension(tmpFileName).ToLower().Equals(".csv"))
            {
                if (!File.Exists(tmpFileName))
                    {
                        throw new Common.KryptonException("Error", Utility.GetCommonMsgVariable("KRYPTONERRCODE0037").Replace("{MSG}", realfilePath));
                    }
                using (GenericParsing.GenericParserAdapter gp = new GenericParsing.GenericParserAdapter(tmpFileName, Encoding.UTF8))
                    {
                        gp.FirstRowHasHeader = true;
                        gp.ColumnDelimiter = ',';
                        ds = gp.GetDataSet();                        
                    }

            }           
            else
            {
                ds = GetExcelDataSet(tmpFileName, dataSetname); 
            }
          }
          finally
           {
               if (File.Exists(tmpFileName))
                    try
                    {
                        File.Delete(tmpFileName);
                    }
                catch{}
           }
            //Validate for unwanted sheets 
          ValidateSheetCount(ds, realfilePath);

          return ds;

        }

        private static void ValidateSheetCount(DataSet ds,string file)
        {
            if (ds.Tables.Count > 1)
                throw new Exception( string.Format("Invalid file error: File contains more than one worksheet. Save only required sheet and remove other invalid sheets in file \"{0}\"", file));
        }

        private static DataSet GetExcelDataSet(string ExcelFilePath, string dataSetname = "dataset")
        {
            DataSet oResultDataSet = new DataSet();
            OledbExcelReader oOledbExcelReader = new OledbExcelReader();
            oResultDataSet = oOledbExcelReader.ReadExcelData(ExcelFilePath);
            if (oResultDataSet == null || oResultDataSet.Tables == null || oResultDataSet.Tables.Count == 0)
            {

                ExcelLibReader oExcelLibReader = new ExcelLibReader();
                oResultDataSet = oExcelLibReader.ReadExcelData(ExcelFilePath);
            }
            for (int i = 0; i < oResultDataSet.Tables.Count; i++)
            {
                while (oResultDataSet.Tables[i].Columns.Count > 300) 
                {
                    oResultDataSet.Tables[i].Columns.RemoveAt(oResultDataSet.Tables[i].Columns.Count - 1);
                }
            }
           int RowsCount= oResultDataSet.Tables[0].Rows.Count;
            return oResultDataSet;
        }
    }
}
