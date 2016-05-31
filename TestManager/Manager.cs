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
using System.IO;
using System.Text;
using Common;
using System.Text.RegularExpressions;
using System.Threading;
using ExcelLib;

namespace Krypton
{
    public class Manager
    {
        private static string _strTestCaseId = string.Empty;
        private static string _strManagerType = string.Empty;

        private static string _testCaseFilePath = string.Empty;
        private static string _testDataFilePath = string.Empty;
        private static string _dbTestDataFilePath = string.Empty;
        private static string _objectRepositoryFilePath = string.Empty;
        private static string _recoverFromPopupFilePath = string.Empty;
        private static string _recoverFromBrowserFilePath = string.Empty;
        private static DirectoryInfo _strResultsSource;
        private static DirectoryInfo _strResultsDestination;

        private static string _resultFilesList = string.Empty;

        public static bool IsTestSuite = false;

        private static string _newTestCaseId = string.Empty;



        public Manager(string testCaseId, string testCasefilename = null)
        {
            char strSeprator = char.Parse(Property.TestCaseIdSeperator);
            _strTestCaseId = testCaseId;
            _newTestCaseId = string.Empty;
            if (testCaseId.StartsWith("TC_"))
            {
                _newTestCaseId = _strTestCaseId.Remove(0, 3); 
            }
            else
            {
                _newTestCaseId = _strTestCaseId;
            }
            SetTestFilesLocation(_strManagerType, testCasefilename);
            Property.CurrentTestCase = _strTestCaseId;
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
                switch (Property.ManagerType.ToLower())
                {
                    case "mstestmanager":
                        TestManager msTestManager = new TestManager();
                        msTestManager.InitExecution();
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
            _strManagerType = managerType;
            char strSeprator = char.Parse(Property.TestCaseIdSeperator);
            switch (_strManagerType.ToLower())
            {
                //For file system, it will get the test files paths from property.cs
                case "filesystem":
                    if (IsTestSuite)
                    {
                        _testCaseFilePath =Path.GetFullPath( Property.TestCaseFilepath + "/" + _strTestCaseId + Property.ExcelSheetExtension);
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(testCasefilename))
                            _testCaseFilePath = Path.GetFullPath(Property.TestCaseFilepath + "/" + _strTestCaseId.Split(strSeprator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension);
                        else
                            _testCaseFilePath = Path.GetFullPath(Property.TestCaseFilepath + "/" + testCasefilename);
                    }
                    if (string.IsNullOrWhiteSpace(testCasefilename))
                        _testDataFilePath = Path.GetFullPath(Property.TestDataFilepath + "/" + "TD_" + _newTestCaseId.Split(strSeprator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension);//New
                    else
                    {
                        _testDataFilePath = Path.GetFullPath(Property.TestDataFilepath + "/" + "TD_" + testCasefilename);
                    }
                    if (string.IsNullOrWhiteSpace(testCasefilename))
                    {
                        _dbTestDataFilePath = Path.GetFullPath(Property.DBTestDataFilepath + "/" + "DB_" + _newTestCaseId.Split(strSeprator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension);
                    }
                    else
                    {
                        _dbTestDataFilePath = Path.GetFullPath(Property.DBTestDataFilepath + "/" + "DB_" + testCasefilename);
                    }
                    _objectRepositoryFilePath = Path.GetFullPath(Property.ObjectRepositoryFilepath + "/" + Property.ObjectRepositoryFilename);
                    _recoverFromPopupFilePath = Path.GetFullPath( Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename);
                    _recoverFromBrowserFilePath = Path.GetFullPath(Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename);
                    break;
                //for QC, it will get the test files paths from the TMQC module.
                case "qc":
                    TestManager objExtManager = new TestManager();
                    //Call DownloadAttachment from TMQC TestManager class to get the path of test files.
                    if (IsTestSuite)
                    {
                        _testCaseFilePath = objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], _strTestCaseId + Property.ExcelSheetExtension);

                    }
                    else
                    {
                        _testCaseFilePath = string.IsNullOrWhiteSpace(testCasefilename) ? objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], Regex.Split(_strTestCaseId, Property.TestCaseIdSeperator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension) : objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], testCasefilename);
                    }

                    _testDataFilePath = string.IsNullOrWhiteSpace(testCasefilename) ? objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], Regex.Split(_strTestCaseId, Property.TestCaseIdSeperator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension) : objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], testCasefilename);

                    _dbTestDataFilePath = objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], Regex.Split(_strTestCaseId, Property.TestCaseIdSeperator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension);
                    _objectRepositoryFilePath = objExtManager.DownloadAttachment(Property.Parameterdic["qcfolder"], Property.ObjectRepositoryFilename);
                    _recoverFromPopupFilePath = Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename;
                    _recoverFromBrowserFilePath = Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename;

                    break;
                case "mstestmanager":
                    // Not Functional yet.
                    if (IsTestSuite)
                    {
                        _testCaseFilePath = Property.TestCaseFilepath + "/" + _strTestCaseId + Property.ExcelSheetExtension;
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(testCasefilename))
                            _testCaseFilePath = Property.TestCaseFilepath + "/" + _strTestCaseId.Split(strSeprator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension;
                        else
                            _testCaseFilePath = Property.TestCaseFilepath + "/" + testCasefilename;
                    }
                    if (string.IsNullOrWhiteSpace(testCasefilename))
                        _testDataFilePath = Property.TestDataFilepath + "/" + "TD_" + _newTestCaseId.Split(strSeprator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension;
                    else
                    {
                        _testDataFilePath = Property.TestDataFilepath + "/" + "TD_" + testCasefilename;
                    }

                    _dbTestDataFilePath = Property.DBTestDataFilepath + "/" + "DB_" + _newTestCaseId.Split(strSeprator)[Property.TestCaseIdParameter] + Property.ExcelSheetExtension;
                    _objectRepositoryFilePath = Property.ObjectRepositoryFilepath + "/" + Property.ObjectRepositoryFilename;
                    _recoverFromPopupFilePath = Property.RecoverFromPopupFilepath + "/" + Property.RecoverFromPopupFilename;
                    _recoverFromBrowserFilePath = Property.RecoverFromBrowserFilePath + "/" + Property.RecoverFromBrowserFilename;
                    break;
            }

        }


        /// <summary>
        /// This method will fetch test steps from Test Case Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Case Steps.</returns>
        public DataSet GetTestCaseXml(string testFlowQuery, bool isGlobalReusabale, bool isSpecificSheet)
        {
            string filePathPath;
            if (isGlobalReusabale)
            {
                string globalActionFileName;
                if (!Path.GetExtension(Utility.GetParameter("reusabledefaultfile").ToLower().Trim()).Equals(Property.ExcelSheetExtension.ToLower()))
                    globalActionFileName = Utility.GetParameter("reusabledefaultfile") + Property.ExcelSheetExtension;
                else
                    globalActionFileName = Utility.GetParameter("reusabledefaultfile").Trim();

                filePathPath = Path.Combine(Property.ReusableLocation, globalActionFileName);

            }
            else if (isSpecificSheet)
            {
                filePathPath = Path.Combine(Property.ReusableLocation, testFlowQuery.Trim() + Property.ExcelSheetExtension);
            }
            else
                filePathPath = _testCaseFilePath;

            //call GetRequiredRows method to read test steps inside test case sheet.
            DataSet strTestCaseFlow = GetDataSet(filePathPath, "testFlow");
            ValidateTestCase(strTestCaseFlow);
            return strTestCaseFlow;
        }


        /// <summary>
        /// Method to Get and Filter the test case id from the SuiteFile.
        /// </summary>
        /// <returns></returns>
        public static DataSet GetSuiteDataSet(string suiteFilePath)
        {
            string datasetName = suiteFilePath.Substring(suiteFilePath.LastIndexOf('\\') + 1);
            suiteFilePath = suiteFilePath + Utility.GetParameter("TestCaseFileExtension");
            var testSuiteDataSet = GetExcelDataSet(suiteFilePath, datasetName);
            return testSuiteDataSet;
        }

        /// <summary>
        /// This method will fetch test data from Test Case Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Data.</returns>
        public DataSet GetTestDataXml()
        {
            var fileName = Property.TestDataSheet.ToLower().Equals("test_data") ? _testDataFilePath : Path.Combine(Path.GetDirectoryName(_testDataFilePath), Property.TestDataSheet + Property.ExcelSheetExtension);
            DataSet strTestDataFlow = GetDataSet(fileName, "TestData"); ;
            ValidateTestData(strTestDataFlow);
            return strTestDataFlow;
        }

        /// <summary>
        /// Get Popup sheet and return as a dataset.
        /// </summary>
        /// <returns></returns>
        public DataSet GetRecoverFromPopupXml()
        {
            DataSet strRecoverFromPopupFlow = GetDataSet(_recoverFromPopupFilePath, "recoveryDataSet");
            ValidateRecoveryFromPopups(strRecoverFromPopupFlow);
            return strRecoverFromPopupFlow;
        }


        /// <summary>
        /// Get Popup sheet and return as a dataset.
        /// </summary>
        /// <returns></returns>
        public DataSet GetRecoverFromBrowserXml()
        {
            DataSet strRecoverFromBrowserFlow = GetDataSet(_recoverFromBrowserFilePath, "recoveryDataSet");
            ValidateRecoveryFromBrowser(strRecoverFromBrowserFlow);
            return strRecoverFromBrowserFlow;
        }

        /// <summary>
        ///This method will fetch DB test data from DBTestData Excel Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Data.</returns>
        public DataSet GetDbTestDataXml()
        {
            DataSet strTestDataFlow = GetDataSet(_dbTestDataFilePath, "TestDbDataflow");
            ValidateDbTestData(strTestDataFlow);
            return strTestDataFlow;
        }

        /// <summary>
        ///  This method will fetch objects from Object Repository Excel Sheet into a dataset
        /// and will return the dataset as XML string.
        /// It will get the object repository sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Object Definition.</returns>
        public static DataSet GetObjectDefinitionXml() //static removed
        {
            DataSet strObjectDefinition = GetDataSet(_objectRepositoryFilePath, "ORDataSet");
            ValidateOr(strObjectDefinition);
            return strObjectDefinition;
        }

        /// <summary>
        ///  This method will copy report files from a temporary location used by Reporting module
        /// to a given location that can be used by the QC or XStudio or similar test management modules. It also
        /// generates a text file containing all result files' path for XStudio.
        /// Source and destination paths will be fetched from the Property.cs file.
        /// </summary>
        /// <returns>True on success else False</returns>
        public static bool UploadTestExecutionResults(List<string> inputTestIds = null)
        {
            bool res;
            try
            {

                _strResultsSource = new DirectoryInfo(Property.ResultsSourcePath).Parent;
                _strResultsDestination = new DirectoryInfo(Property.ResultsDestinationPath);

                // Copy all result files from source directory to destination directory.
                res = CopyAll(_strResultsSource, _strResultsDestination);

                // Create a temporary file to make sure that the copy process has been completed.
                bool resFinished = true;
                if (_strResultsDestination.Parent != null)
                {
                    StreamWriter sw = new StreamWriter(_strResultsDestination.Parent.FullName + "/ResultCopyCompleted.txt");
                    sw.WriteLine("Key: Copy Process Completed!");
                    sw.Close();
                }

                var directoryInfo = new DirectoryInfo(Property.ResultsSourcePath).Parent;
                if (directoryInfo != null)
                {
                    DirectoryInfo drParent = directoryInfo.Parent;

                    //to check if user wants to keep report history
                    string fileExt = string.Empty;
                    if (_strResultsDestination.Name.LastIndexOf('-') >= 0 && Utility.GetParameter("KeepReportHistory").Equals("true", StringComparison.OrdinalIgnoreCase) == true)
                        fileExt = _strResultsDestination.Name.Substring(_strResultsDestination.Name.LastIndexOf('-'),
                            _strResultsDestination.Name.Length - _strResultsDestination.Name.LastIndexOf('-'));

                    string rcProcessId = string.Empty;
                    if (!string.IsNullOrWhiteSpace(Property.RcProcessId))
                        rcProcessId = "-" + Property.RcProcessId; //Adding process id for Remote execution

                    if (drParent != null)
                        foreach (FileInfo fi in drParent.GetFiles())
                        {
                            if (fi.Name.IndexOf("HtmlReport", StringComparison.OrdinalIgnoreCase) >= 0
                                || fi.Name.IndexOf("logo", StringComparison.OrdinalIgnoreCase) >= 0
                                || fi.Name.IndexOf(Property.ReportZipFileName, StringComparison.OrdinalIgnoreCase) >= 0
                                || fi.Name.IndexOf("." + Property.ScriptLanguage, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (fi.Name.IndexOf("mail.html", StringComparison.OrdinalIgnoreCase) >= 0) // no need to copy this file to krypton result folder while sending mail through krypton
                                    continue;
                                string fileName;
                                if (fi.Name.IndexOf("logo", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    fileName = fi.Name;
                                }
                                else if (fi.Name.IndexOf("HtmlReport", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    string htmlFileExt;
                                    if (!string.IsNullOrWhiteSpace(Utility.GetParameter("TestSuite")))
                                    {
                                        htmlFileExt = Utility.GetParameter("TestSuite");
                                    }
                                    else
                                    {
                                        string[] testCaseIds = Utility.GetParameter("TestCaseId").Split(',');
                                        if (testCaseIds.Length > 1) htmlFileExt = testCaseIds[0] + "...multiple";
                                        else htmlFileExt = Utility.GetParameter("TestCaseId");
                                    }

                                    htmlFileExt = htmlFileExt + rcProcessId;

                                    if (fi.Name.IndexOf("HtmlReports", StringComparison.OrdinalIgnoreCase) >= 0)
                                        fileName = fi.Name.Substring(0, fi.Name.LastIndexOf('.')) +
                                                   "-" + htmlFileExt + fileExt + "s" + fi.Name.Substring(fi.Name.LastIndexOf('.'), fi.Name.Length - fi.Name.LastIndexOf('.'));
                                    else
                                        fileName = fi.Name.Substring(0, fi.Name.LastIndexOf('.')) +
                                                   "-" + htmlFileExt + fileExt + fi.Name.Substring(fi.Name.LastIndexOf('.'), fi.Name.Length - fi.Name.LastIndexOf('.'));
                                }
                                else
                                {
                                    fileName = fi.Name.Substring(0, fi.Name.LastIndexOf('.')) +
                                               fileExt +
                                               fi.Name.Substring(fi.Name.LastIndexOf('.'), fi.Name.Length - fi.Name.LastIndexOf('.'));
                                }


                                if (_strResultsDestination.Parent != null)
                                {
                                    fi.CopyTo(_strResultsDestination.Parent.FullName + "/" + fileName, true);

                                    //if user wants to keep report history, folder name will be different, so need to change links in the html file also
                                    if (string.IsNullOrWhiteSpace(fileExt) == false && fi.Name.Contains("HtmlReport"))
                                    {
                                        string htmlStr = File.ReadAllText(_strResultsDestination.Parent.FullName + "/" + fileName, Encoding.UTF8);

                                        string testCaseId = string.Empty;

                                        //Handle Test Case Id conditions
                                        #region check for test suite
                                        string testSuite = Utility.GetParameter("TestSuite");
                                        if (!string.IsNullOrWhiteSpace(testSuite.Trim()) && string.IsNullOrWhiteSpace(Utility.GetParameter("TestCaseId").Trim()))
                                        {
                                            SetTestFilesLocation(Property.ManagerType); // get ManagerType from property.cs
                                            Manager objTestManagerSuite = new Manager(testSuite);
                                            var testSuiteData = objTestManagerSuite.GetTestCaseXml(Property.TestQuery, false, false);
                                            for (int testCaseCnt = 0; testCaseCnt < testSuiteData.Tables[0].Rows.Count; testCaseCnt++)
                                            {
                                                if (!string.IsNullOrWhiteSpace(testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString().Trim()))
                                                {
                                                    if (string.IsNullOrWhiteSpace(testCaseId))
                                                        testCaseId = testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID].ToString();
                                                    else
                                                        testCaseId = testCaseId + "," +
                                                                     testSuiteData.Tables[0].Rows[testCaseCnt][KryptonConstants.TEST_CASE_ID];
                                                }
                                            }

                                        }
                                        else
                                        {
                                            testCaseId = Utility.GetParameter("TestCaseId");//override testcaseid parameter over test suite
                                        }

                                        #endregion

                                        var testCases = inputTestIds != null ? inputTestIds.ToArray() : testCaseId.Split(',');

                                        string[] testCasesIds = new string[testCases.Length];

                                        int k = 0;
                                        foreach (string testCase in testCases)
                                        {
                                            if (string.Join(";", testCasesIds).Contains(testCase + ";") == false)
                                            {
                                                testCasesIds[k] = testCase;
                                                htmlStr = htmlStr.Replace("href='" + testCase + "\\",
                                                    "href='" + testCase + rcProcessId + fileExt + "\\");
                                                k++;
                                            }
                                        }
                                        File.WriteAllText(_strResultsDestination.Parent.FullName + "/" + fileName, htmlStr, Encoding.UTF8);

                                    }
                                }
                            }
                        }
                }
                int resCheckLimit = 0;
                do
                {
                    if (_strResultsDestination.Parent != null && File.Exists(_strResultsDestination.Parent.FullName + "/ResultCopyCompleted.txt"))
                    {
                        File.Delete(_strResultsDestination.Parent.FullName + "/ResultCopyCompleted.txt");
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
                throw new KryptonException("error", exception.Message);
            }

            var fileUploadPath = Property.ResultsDestinationPath + "/" + (new DirectoryInfo(Property.ResultsSourcePath).Name);

            // Extra actions are performed for particular test manager.
            switch (_strManagerType.ToLower())
            {
                case "qc":
                    // Not Functional yet
                    break;
                case "xstudio":
                    // XStudio requires a trigger that tells about the end of the execution.
                    // Generation of test_completed.txt file will notify XStudio.
                    TestManager.UploadTestResults(_resultFilesList, fileUploadPath);
                    break;

                case "mstestmanager":
                    // Not Functional yet
                    TestManager.UploadTestResults(Property.FinalXmlPath, Property.FinalExecutionStatus);
                    break;
            }

            return res;
        }


        /// <summary>
        /// This method copies the content of a given directory into a destination directory.
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
                _resultFilesList = string.Empty;
                //copy all the files into source directory to target directory.
                foreach (FileInfo fi in source.GetFiles())
                {
                    fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);

                    //this string stores the filenames along with full path for XStudio test_completed.txt file.
                    _resultFilesList += (fi.Name.Split('.')[1] == "xml") ? "Result_File=" + Path.Combine(target.ToString(), fi.Name) + "\r\n" : "Upload_File=" + Path.Combine(target.ToString(), fi.Name) + "\r\n";
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
                throw new KryptonException("error", exception.Message);
            }

            return true;
        }

        /// <summary>
        ///  Validate the column names for test case dataset.
        /// </summary>
        /// <param name="testCaseXml">DataSet of the current test case sheet to be executed</param>
        public static void ValidateTestCase(DataSet testCaseXml)
        {
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
                    var colFound = false;
                    for (int j = 0; j < testCaseXml.Tables[0].Columns.Count; j++)
                    {
                        if (testCaseColumns[i] == testCaseXml.Tables[0].Columns[j].ColumnName.ToLower())
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
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0031").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for test data dataset.
        /// </summary>
        /// <param name="testDataXml">DataSet of the current test data sheet</param>
        public static void ValidateTestData(DataSet testDataXml)
        {
            string[] testDataColumns = { "row_no" };

            // Verify each column name in test data sheet.
            try
            {
                for (int i = 0; i < testDataColumns.GetLength(0); i++)
                {
                    var colFound = false;
                    for (int j = 0; j < testDataXml.Tables[0].Columns.Count; j++)
                    {
                        if (testDataColumns[i] == testDataXml.Tables[0].Columns[j].ColumnName.ToLower())
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
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0032").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for object repository dataset.
        /// </summary>
        /// <param name="objectRepositoryXml">DataSet of the current object repository sheet</param>
        public static void ValidateOr(DataSet objectRepositoryXml)
        {
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
                    var colFound = false;
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
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0033").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for db test data dataset.
        /// </summary>
        /// <param name="dbTestDataXml">DataSet of the current db test data sheet</param>
        public static void ValidateDbTestData(DataSet dbTestDataXml)
        {
            string[] dbTestDataColumns = { "sno",
                                           "sqlquery" };

            // Verify each column name in test data sheet.
            try
            {
                for (int i = 0; i < dbTestDataColumns.GetLength(0); i++)
                {
                    var colFound = false;
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
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0034").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for recovery from popup dataset.
        /// </summary>
        /// <param name="recoveryFromPopupsXml"></param>
        public static void ValidateRecoveryFromPopups(DataSet recoveryFromPopupsXml)
        {
            string[] recoveryFromPopupsColumns = { "popuptext",
                                                   "action" };

            // Verify each column name in PopUp Recovery sheet.
            try
            {
                for (int i = 0; i < recoveryFromPopupsColumns.GetLength(0); i++)
                {
                    var colFound = false;
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
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0035").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  Validate the column names for recovery from browser dataset.
        /// </summary>
        /// <param name="recoveryFromBrowserXml"></param>
        public static void ValidateRecoveryFromBrowser(DataSet recoveryFromBrowserXml)
        {
            string[] recoveryFromBrowserColumns = { "recovery_keyword",
                                                    "recovery_details",
                                                    "action" };

            // Verify each column name in Browser Recovery sheet.
            try
            {
                for (int i = 0; i < recoveryFromBrowserColumns.GetLength(0); i++)
                {
                    var colFound = false;
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
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0036").Replace("{MSG}", e.Message));
            }

        }

        /// <summary>
        ///  To generate a temp file and return file path
        /// </summary>
        /// <param name="oriFilePath"></param>
        /// <returns></returns>
        private static string GetTempFile(string oriFilePath)
        {
            var tmpFileName = Utility.GetTemporaryFile(Path.GetExtension(oriFilePath));
            if (File.Exists(oriFilePath))
            {
                File.Copy(oriFilePath, tmpFileName);
            }
            else
            {
                throw new KryptonException("Error", Utility.GetCommonMsgVariable("KRYPTONERRCODE0037").Replace("{MSG}", oriFilePath));
            }
            return tmpFileName;
        }

        /// <summary>
        ///  This method read an csv file and returns a dataset.
        /// </summary>
        /// <param name="realfilePath"></param>
        /// <param name="dataSetname"></param>
        /// <returns></returns>
        private static DataSet GetDataSet(string realfilePath, string dataSetname = "dataset")
        {
            DataSet ds = new DataSet(dataSetname);
            string tmpFileName = GetTempFile(realfilePath);
            try
            {
                var extension = Path.GetExtension(tmpFileName);
                if (extension != null && extension.ToLower().Equals(".csv"))
                {
                    if (!File.Exists(tmpFileName))
                    {
                        throw new KryptonException("Error", Utility.GetCommonMsgVariable("KRYPTONERRCODE0037").Replace("{MSG}", realfilePath));
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
                    catch
                    {
                        // ignored
                    }
            }
            //Validate for unwanted sheets 
            ValidateSheetCount(ds, realfilePath);

            return ds;

        }

        private static void ValidateSheetCount(DataSet ds, string file)
        {
            if (ds.Tables.Count > 1)
                throw new Exception(string.Format("Invalid file error: File contains more than one worksheet. Save only required sheet and remove other invalid sheets in file \"{0}\"", file));
        }

        private static DataSet GetExcelDataSet(string excelFilePath, string dataSetname = "dataset")
        {
            OledbExcelReader oOledbExcelReader = new OledbExcelReader();
            var oResultDataSet = oOledbExcelReader.ReadExcelData(excelFilePath);
            if (oResultDataSet == null || oResultDataSet.Tables.Count == 0)
            {

                ExcelLibReader oExcelLibReader = new ExcelLibReader();
                oResultDataSet = oExcelLibReader.ReadExcelData(excelFilePath);
            }
            for (int i = 0; i < oResultDataSet.Tables.Count; i++)
            {
                while (oResultDataSet.Tables[i].Columns.Count > 300)
                {
                    oResultDataSet.Tables[i].Columns.RemoveAt(oResultDataSet.Tables[i].Columns.Count - 1);
                }
            }
            return oResultDataSet;
        }
    }
}
