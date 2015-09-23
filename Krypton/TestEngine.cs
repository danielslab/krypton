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
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Common;
using Reporting;
using System.Windows.Forms;
using System.Diagnostics;
using System.ComponentModel;
using System.Threading;
using System.Drawing;
using System.Configuration;

namespace Krypton
{
    class TestEngine : Variables
    {

        #region Member Variables
        private static DataSet ORTestData = null; 
        private static int stepNos = 0; 
        private static int failureCount = 0;
        private static DataSet xmlTestDataSet = null; 
        private static Krypton.Manager objTestManager = null; 
        private static string cleanupTestCase = string.Empty; 
        public static IKryptonLogger logwriter = null;
        public static bool executeThread = false;
        #endregion


        /// <summary>
        /// This is Main console engine
        /// execute test cases and generate log file
        /// </summary>
        /// <param name="args">console argumants which will override parameter.ini default parameters</param> 
        /// <returns></returns>
        static int Main(string[] args)
        {
            string Version = ConfigurationSettings.AppSettings["Version"];
            Console.Title = "Krypton" + " v" + "2.0.0";
            try
            {
                //To find application startup path.
                string applicationPath = Application.StartupPath;
                if (applicationPath.Length > 0)
                {

                    Common.Property.ApplicationPath = applicationPath + "\\";
                }
                else
                {
                    DirectoryInfo dr = new DirectoryInfo("./");
                    Common.Property.ApplicationPath = dr.FullName;
                }

                initializeIniPath(args);
                Common.Utility.SetParameter("ApplicationPath", Common.Property.ApplicationPath);
                Common.Utility.SetVariable("ApplicationPath", Common.Property.ApplicationPath);

                Common.Utility.GetCommonMessageData();

                //Parse ini file and put all parameters into global dictionary

                #region Adding internal parameter, runbyevents.
                /* Parameter "runbyevents" means, certain methods will run using browser events instead of using native driver implemented methods.
                   This can also be passed from command line arguments, but is not added in parameter.ini file.
                   When its value is true, browser events will be used to perform operations instead of native methods.
                   E.g. arguments[0].click(); will be used instead of testobject.click when this parameter is true
                */
                Common.Property.parameterdic.Add("runbyevents", "false");
                Common.Property.runtimedic.Add("runbyevents", "false");
                Common.Utility.SetParameter("keyword", string.Empty);
                Common.Utility.SetVariable("keyword", string.Empty);
                Common.Utility.SetParameter("currentRelease", "0");
                Common.Utility.SetVariable("currentRelease", "0");
                if (Path.IsPathRooted(Property.DESTINATION_FOLDER_DOWNLOAD))
                {
                    Common.Utility.SetParameter("downloadpath", Property.DESTINATION_FOLDER_DOWNLOAD);
                    Common.Utility.SetVariable("downloadpath", Property.DESTINATION_FOLDER_DOWNLOAD);
                }
                else
                {
                    Common.Utility.SetParameter("downloadpath", string.Concat(Common.Property.ApplicationPath, Property.DESTINATION_FOLDER_DOWNLOAD));
                    Common.Utility.SetVariable("downloadpath", string.Concat(Common.Property.ApplicationPath, Property.DESTINATION_FOLDER_DOWNLOAD));
                }
                #endregion

                Common.Utility.CollectkeyValuePairs();

                #region Parse arguments passed to executable. This may over-write ini file parameters, if names are same
                if (args.Length != 0)
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        string[] KeyValuePair = args[i].Trim().Split('=');
                        string Key = KeyValuePair[0].Trim().ToLower();
                        string Value = KeyValuePair[1].Trim();
                        Common.Utility.SetParameter(Key, Value);
                        Common.Utility.SetVariable(Key, Value);
                    }
                }
                #endregion
                // Set Environment File Paths
                if (Path.IsPathRooted(Common.Utility.GetParameter("EnvironmentFileLocation")))  
                    Common.Property.EnvironmentFileLocation = Common.Utility.GetParameter("EnvironmentFileLocation");
                else
                    Common.Property.EnvironmentFileLocation = Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("EnvironmentFileLocation"));

                Common.Utility.UpdateKeyValuePairs();
                Common.Utility.ValidateParameters();
                Common.Property.RECOVERY_COUNT = Convert.ToInt16(Utility.GetParameter("recoverycount"));
                // modified to set the value only if it is given in parameters.ini
                if (Utility.GetParameter("globaltimeout") != string.Empty)
                    Common.Property.GlobalTimeOut = Utility.GetParameter("globaltimeout");


                //In case manager type is MSTestManager, user can pass on any of the following options:
                //MSTestManager, MSTM, MTM, VSTM, VSTestManager
                if (Common.Utility.GetParameter("ManagerType").StartsWith("MSTestManager", StringComparison.OrdinalIgnoreCase) ||
                    Common.Utility.GetParameter("ManagerType").StartsWith("MSTM", StringComparison.OrdinalIgnoreCase) ||
                    Common.Utility.GetParameter("ManagerType").StartsWith("MTM", StringComparison.OrdinalIgnoreCase) ||
                    Common.Utility.GetParameter("ManagerType").StartsWith("VSTM", StringComparison.OrdinalIgnoreCase) ||
                    Common.Utility.GetParameter("ManagerType").StartsWith("VSTestManager", StringComparison.OrdinalIgnoreCase))
                {
                    Common.Utility.SetParameter("ManagerType", "MSTestManager");
                    Common.Utility.SetVariable("ManagerType", "MSTestManager");
                }


                //Get browser settings
                Common.Utility.InitializeBrowser(Common.Utility.GetParameter(Common.Property.BrowserString));

                Common.Property.ApplicationURL = Common.Utility.GetParameter("ApplicationURL");

                int num;
                if (Common.Property.parameterdic.ContainsKey("waitforalert"))
                    if (!string.IsNullOrEmpty(Common.Utility.GetParameter("Waitforalert")) && int.TryParse(Common.Utility.GetParameter("Waitforalert"), out num))
                        Common.Property.Waitforalert = int.Parse(Common.Utility.GetParameter("Waitforalert"));

                Common.Property.ExcelSheetExtension = Common.Utility.GetParameter("TestCaseFileExtension");
                if (!Property.ExcelSheetExtension.StartsWith("."))
                {
                    Property.ExcelSheetExtension = "." + Property.ExcelSheetExtension;
                }

                //Get Test Case File path
                if (Path.IsPathRooted(Common.Utility.GetParameter("TestCaseLocation")))
                    Common.Property.TestCaseFilepath = Common.Utility.GetParameter("TestCaseLocation");
                else
                    Common.Property.TestCaseFilepath = Path.Combine(Common.Property.IniPath, Common.Utility.GetParameter("TestCaseLocation"));

                //set the ManagerType
                if (Common.Utility.GetParameter("ManagerType").IsNullOrWhiteSpace() == false)
                    Common.Property.ManagerType = Common.Utility.GetParameter("ManagerType");

                if (Common.Utility.GetParameter("ListOfUniqueCharacters").IsNullOrWhiteSpace() == false)
                {
                    Common.Property.ListOfUniqueCharacters = Common.Utility.GetParameter("ListOfUniqueCharacters").Split(',');
                }

                //Handle Test Case Id conditions
                #region check for test suite
                if (Common.Utility.GetParameter("TestSuite").Trim().IsNullOrWhiteSpace() && Common.Utility.GetParameter("TestSuite") != string.Empty)
                {
                    Common.Utility.SetParameter("TestSuite", Common.Utility.GetParameter("TestCaseId").Trim().Split('.')[0]);
                }
                else
                {
                    testSuite = Common.Utility.GetParameter("TestSuite").Trim();
                }
                if (Common.Utility.GetParameter("TestSuite")!=string.Empty)
                {
                    string testsuiteFilePath = Path.GetFullPath(Common.Property.IniPath + "/" + Common.Utility.GetParameter("TestSuiteLocation") + "/" + Common.Utility.GetParameter("TestSuite"));
                    Common.Utility.SetParameter("TestCaseId", Common.Utility.GetTestCaseIdFromSuiteFile(testsuiteFilePath));
                }
                if (!Common.Utility.GetParameter("TestCaseId").Trim().IsNullOrWhiteSpace())
                    testSuite = Common.Utility.GetParameter("TestCaseId").Trim().Split('.')[0];

                testCaseId = KryptonUtility.getTestCaseIDs();

                Krypton.Manager.isTestSuite = false;

                //Test Cases that Input From the Parameter.ini
                string[] inputTestCases = Common.Property.runtimedic["testcaseid"].ToString().Split(',');
                #endregion

                //Test Cases that are exit in file
                string[] matchingTestCase = testCaseId.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                string finalTestCases = null;

                //Loop to Check Repitative Test Cases
                for (int i = 0; i < matchingTestCase.Length; i++)
                {
                    for (int j = 0; j < inputTestCases.Length; j++)
                    {
                        if (matchingTestCase[i].Equals(inputTestCases[j], StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (finalTestCases.IsNullOrWhiteSpace())
                                finalTestCases = inputTestCases[j];
                            else
                                finalTestCases = finalTestCases + "," + inputTestCases[j];
                        }
                    }
                }

                if (Common.Utility.GetParameter("TestSuite") != string.Empty)
                {
                    foreach (var item in matchingTestCase)
                    {
                        finalTestCases = finalTestCases + "," + item;
                    }
                    finalTestCases = finalTestCases.Substring(1);
                }

                testCases = finalTestCases.Split(',');
                // Check if there is at least one test case id for test execution
                if (testCases.Length.Equals(0))
                {
                    Console.WriteLine(new Common.KryptonException(exceptions.ERROR_NOTESTCASEFOUND));
                    return testSuiteResult;
                }

                KryptonUtility.SetProjectFilesPaths();
                #region to get database from different ini files
                string[] sqlConnectionStr = Common.Property.DbConnectionString.Trim().Split(';');

                for (int sqlParamCnt = 0; sqlParamCnt < sqlConnectionStr.Length; sqlParamCnt++)
                {
                    if (sqlConnectionStr[sqlParamCnt].Trim().IndexOf("database", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string oriDatabase = sqlConnectionStr[sqlParamCnt];
                        if (Common.Utility.GetParameter("write").IsNullOrWhiteSpace() == false)
                        {
                            sqlConnectionStr[sqlParamCnt] = "database=" + Common.Utility.GetParameter("write");
                        }
                        else
                        {
                            sqlConnectionStr[sqlParamCnt] = oriDatabase;
                        }

                        Common.Property.SqlConnectionStringWrite = string.Join(";", sqlConnectionStr);
                        if (Common.Utility.GetParameter("read").IsNullOrWhiteSpace() == false)
                        {
                            sqlConnectionStr[sqlParamCnt] = "database=" + Common.Utility.GetParameter("read");
                        }
                        else
                        {
                            sqlConnectionStr[sqlParamCnt] = oriDatabase;
                        }
                        Common.Property.SqlConnectionStringRead = string.Join(";", sqlConnectionStr);
                    }
                }
                #endregion

                //dbtestdata file path
                if (Path.IsPathRooted(Common.Utility.GetParameter("DBQueryFilePath")))
                    Common.Property.SqlQueryFilePath = Common.Utility.GetParameter("DBQueryFilePath");
                else
                    Common.Property.SqlQueryFilePath = string.Concat(Common.Property.IniPath, Common.Utility.GetParameter("DBQueryFilePath"));

                //to split testcaseid for test case excel file
                Common.Property.TestCaseIDSeperator = Common.Utility.GetParameter("TestCaseIDSeperator");

                //to fetch test case excel file from testcaseid
                Common.Property.TestCaseIDParameter = int.Parse(Common.Utility.GetParameter("TestCaseIDParameter"));

                Common.Property.DebugMode = Common.Utility.GetParameter("debugmode").Trim();

                //for snapshot option
                Common.Property.SnapshotOption = Common.Utility.GetParameter("SnapshotOption");

                //for remote execution initialization
                Common.Property.IsRemoteExecution = Common.Utility.GetParameter("RunRemoteExecution");

                //Force snapshot option in case of remote driver
                if (Common.Property.IsRemoteExecution.Equals("true") && Common.Property.SnapshotOption.Equals("always"))
                {
                    Common.Property.SnapshotOption = "on page change";
                }

                Common.Property.RemoteUrl = Common.Utility.GetParameter("RunOnRemoteBrowserUrl");

                //for email templates
                if (Path.IsPathRooted(Common.Utility.GetParameter("EmailStartTemplate")))
                    Common.Property.EmailStartTemplate = Common.Utility.GetParameter("EmailStartTemplate");
                else
                    Common.Property.EmailStartTemplate = string.Concat(Common.Property.IniPath, Common.Utility.GetParameter("EmailStartTemplate"));

                if (Path.IsPathRooted(Common.Utility.GetParameter("EmailEndTemplate")))
                    Common.Property.EmailEndTemplate = Common.Utility.GetParameter("EmailEndTemplate");
                else
                    Common.Property.EmailEndTemplate = string.Concat(Common.Property.IniPath, Common.Utility.GetParameter("EmailEndTemplate"));


                //scripting language
                if (!Common.Utility.GetVariable("ScriptLanguage").IsNullOrWhiteSpace())
                    Common.Property.ScriptLanguage = Common.Utility.GetVariable("ScriptLanguage");
            }
            catch (Exception exception)
            {
                Console.WriteLine(new Common.KryptonException(exception.Message));
                return testSuiteResult;
            }


            if (!Common.Utility.GetVariable("RCProcessId").IsNullOrWhiteSpace() && !Common.Utility.GetVariable("RCProcessId").Equals("RCProcessId", StringComparison.OrdinalIgnoreCase))
                Common.Property.RCProcessId = Common.Utility.GetVariable("RCProcessId");

            if (!Common.Utility.GetVariable("RCMachineId").IsNullOrWhiteSpace() && !Common.Utility.GetVariable("RCMachineId").Equals("RCMachineId", StringComparison.OrdinalIgnoreCase))
            {
                Common.Property.RCMachineId = Common.Utility.GetVariable("RCMachineId");

                //Update saucelabs machine status to protect identity of username and api key
                if (Common.Property.RCMachineId.ToLower().Contains(KryptonConstants.BROWSER_SAUCELABS))
                    Common.Property.RCMachineId = KryptonConstants.BROWSER_SAUCELABS;
            }

            else
            {
                Common.Property.RCMachineId = Environment.MachineName;
                Common.Utility.SetVariable("RCMachineId", Environment.MachineName);
                Common.Utility.SetParameter("RCMachineId", Environment.MachineName);
            }

            if (!Environment.UserName.IsNullOrWhiteSpace())
                Common.Property.RCUserName = Environment.UserName; //set logged in username


            if (Common.Utility.GetParameter("FailedCountForExit").IsNullOrWhiteSpace())
            {
                Common.Utility.SetParameter("FailedCountForExit", Common.Property.FailedCountForExit);
            }

            if (Common.Utility.GetParameter("StartParallelRecovery").IsNullOrWhiteSpace())
            {
                Common.Utility.SetParameter("StartParallelRecovery", Common.Property.StartParallelRecovery);
            }

            Common.Property.DATE_TIME = Common.Utility.GetParameter("DateTimeFormat").Replace("/", "\\/");

            //Execution start date and time set
            DateTime dtNow = DateTime.Now;

            #region source path location and clear temp folder
            string sourcePath = string.Empty;
            sourcePath = "KRYPTONResults-" + Guid.NewGuid().ToString();

            //HTML File location
            Common.Property.HtmlFileLocation = string.Format("{0}{1}", Path.GetTempPath(), sourcePath);
            Property.listOfFilesInTempFolder.Add(Property.HtmlFileLocation);
            #endregion

            Common.Property.ExecutionStartDateTime = dtNow.ToString(Common.Property.DATE_TIME);
            filePath = new string[testCases.Length];

            bool validation = false; //need to capture the validation steps so that it will not check for each test case

            //when running in batch mode (more than one test case or in test suite), value of this parameter is always true, irrespective of what user specified
            if (testCases.Length > 1) Utility.SetParameter("closebrowseroncompletion", "true");

            // final test case id to be stored in final test case variable
            Property.FinalTestCase = testCases[testCases.Length - 1];


            #region  Setup Test Environment on Local Workstation
            for (int testCaseCnt = 0; testCaseCnt < testCases.Length; testCaseCnt++)
            {
                // This section handles setting test environment using command line provided in Environment specific ini files
                // Retrieve command to be executed from variables
                string envSetupCommand = Utility.GetVariable("EnvironmentSetupBatch");
                envSetupCommand = Utility.ReplaceVariablesInString(envSetupCommand);

                // Do not attempt to setup environment if no script needs to be executed
                if (!(envSetupCommand.IsNullOrWhiteSpace()))
                {
                    if (File.Exists(Path.Combine(Common.Property.EnvironmentFileLocation, envSetupCommand)))
                    {
                        Process setupProcess = new Process();
                        setupProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        setupProcess.StartInfo.UseShellExecute = false;
                        setupProcess.StartInfo.FileName = Path.Combine(Common.Property.EnvironmentFileLocation, envSetupCommand);
                        setupProcess.StartInfo.ErrorDialog = false;

                        //Pass on couple of extra arguments to setup batch file
                        //These includes BrowserType, TestEnvironment
                        setupProcess.StartInfo.Arguments = Utility.GetParameter("Browser") +
                                                           " " + Utility.GetParameter("Environment") +
                                                           " " + "\"" + Common.Utility.GetParameter("TestSuite").Trim() + "\"" +
                                                           " " + "\"" + Common.Property.RCProcessId + "\"";

                        // Start the process
                        setupProcess.Start();
                        setupProcess.WaitForExit();
                        try
                        {
                            if (!setupProcess.HasExited)
                            {
                                setupProcess.Kill();
                            }
                        }
                        catch (Exception e)
                        {
                            Common.KryptonException.writeexception(e);
                        }
                    }
                    else
                    {
                        Console.WriteLine(new Common.KryptonException("Environment file: " + Common.Property.ApplicationPath + envSetupCommand + " is not exist."));
                        return testSuiteResult;
                    }
                }
                #endregion


                Common.Property.ResultsSourcePath = string.Format("{0}{1}\\{2}\\{3}", Path.GetTempPath(), sourcePath, testCases[testCaseCnt].Trim(),
                                                                       (testCaseCnt + 1).ToString());


                //if user wants to keep report history, it will be available with dateandtime appended in the log destination folder
                string logDestinationExt = string.Empty;
                if (Common.Utility.GetParameter("KeepReportHistory").Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    logDestinationExt = "-" + dtNow.ToString("ddMMyyhhmmss");

                }
                if (!Common.Property.RCProcessId.IsNullOrWhiteSpace())
                    logDestinationExt = "-" + Common.Property.RCProcessId + logDestinationExt;

                if (Path.IsPathRooted(Common.Utility.GetParameter("LogDestinationFolder")))
                    Common.Property.ResultsDestinationPath = string.Format("{0}\\{1}{2}", Common.Utility.GetParameter("LogDestinationFolder"),
                                                                           testCases[testCaseCnt].Trim(), logDestinationExt);
                else
                    Common.Property.ResultsDestinationPath = string.Format("{0}\\{1}{2}", string.Concat(Common.Property.IniPath,
                                                                    Common.Utility.GetParameter("LogDestinationFolder")),
                                                                           testCases[testCaseCnt].Trim(), logDestinationExt);

                //Initialize the logwriter to record any message/exception after test step completion
                try
                {
                    logwriter = new KryptonFileLogWriter(testCaseCnt);
                }
                catch (Exception e)
                {
                    Common.KryptonException.writeexception(e);
                    Console.WriteLine(Common.exceptions.ERROR_INVALIDTESTID + testCases[testCaseCnt].ToString() + "in Test Suite: " + testSuite); // If test case Id will be of more than 255 characters.. 
                    return testSuiteResult;
                }

                #region Creating the xml log file path
                try
                {
                    filePath[testCaseCnt] = string.Format("{0}\\{1}", Common.Property.ResultsSourcePath, Common.Property.LogFileName);
                    xmlLog = new Reporting.LogFile(filePath[testCaseCnt]); //xml log file initialization
                    xmlLog.AddTestAttribute("TestCase Id", testCases[testCaseCnt].Trim());
                    if (!Common.Property.RCMachineId.IsNullOrWhiteSpace())
                        xmlLog.AddTestAttribute("RCMachineId", Common.Property.RCMachineId);

                }
                catch (Exception exception)
                {
                    Console.WriteLine(new Common.KryptonException(exception.Message));
                    return testSuiteResult;
                }
                #endregion

                Common.Property.ValidateSetup = Common.Utility.GetParameter("ValidateSetup");

                if (Common.Property.ValidateSetup.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    #region validation

                    //once it validates it will not run for each test case
                    try
                    {
                        if (validation != true)
                        {
                            Common.Validate.getDriverValidation(
                                Common.Utility.GetParameter(Common.Property.DRIVER_STRING).ToLower());
                            Common.Property.Status = Common.Validate.validate.validationProcess();
                            if (Common.Property.Status == Common.ExecutionStatus.Fail)
                            {
                                ReportHandler.WriteExceptionLog(new Common.KryptonException("Validation Faild"), 0, "Validation");
                                return testSuiteResult;
                            }
                            else validation = true;
                        }

                    }
                    catch (Exception exception)
                    {
                        ReportHandler.WriteExceptionLog(new KryptonException(exception.Message), 0, "Validation");
                        return testSuiteResult;
                    }

                    #endregion
                }
                else
                {
                    validation = true;
                }
                //Added code for focus mouse to origin
                if (testCaseCnt == 0 && Property.IsRemoteExecution.ToLower().Equals("false"))
                {

                    System.Windows.Forms.Cursor.Position = new Point(0, 0);
                }

                objTestManager = new Krypton.Manager(testCases[testCaseCnt].Trim());

                //Call InitExecution method to allow manager pass on control 
                //to respective test manager for required settings
                try
                {
                    Krypton.Manager.InitTestCaseExecution();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                #region Fetch OR data
                ////if OR data is fetched, this step will not execute for each test case
                DataSet xmlRecoverFromPopupFlow = new DataSet();
                DataSet xmlRecoverFromBrowserFlow = new DataSet();
                try
                {
                    if (ORTestData == null)
                    {

                        if (!Path.GetExtension(Utility.GetParameter("ORFileName")).ToLower().Equals(Property.ExcelSheetExtension.ToLower()))
                            Property.ObjectRepositoryFilename = Utility.GetParameter("ORFileName").ToLower() + Property.ExcelSheetExtension;
                        else
                            Property.ObjectRepositoryFilename = Utility.GetParameter("ORFileName").ToLower();
                        Krypton.Manager.SetTestFilesLocation(Common.Property.ManagerType);  // get ManagerType from property.cs
                        OrData(); //To Fetch object Repository data and set in private global dataset

                        if (testCaseCnt == 0 && !Common.Utility.GetParameter("RunRemoteExecution").ToLower().Equals("true"))
                        {
                            //reading RecoverPopup File.
                            xmlRecoverFromPopupFlow = objTestManager.GetRecoverFromPopupXml();
                            xmlRecoverFromBrowserFlow = objTestManager.GetRecoverFromBrowserXml();

                            #region Start Exe for Popups handling.
                            try
                            {
                                if (string.Compare(Common.Utility.GetVariable("StartParallelRecovery"), "true", true) == 0)
                                {
                                    string ext = Path.GetExtension(Property.ParallelRecoverySheetName);
                                    if (!(ext.Length > 0 && ext != null))
                                        Property.ParallelRecoverySheetName = Property.ParallelRecoverySheetName + Property.ExcelSheetExtension;
                                    if (!File.Exists(Path.Combine(Property.RecoverFromPopupFilepath, Property.ParallelRecoverySheetName)))
                                    {
                                        Console.WriteLine(ConsoleMessages.MSG_PARALLEL_RECOVERY);
                                        Console.WriteLine("Please make sure  " + Property.ParallelRecoverySheetName + "  exist at  " + Property.RecoverFromPopupFilepath + "  directory.");
                                        return testSuiteResult;
                                    }
                                    ProcessStartInfo processInfo = new ProcessStartInfo();
                                    processInfo.FileName = Common.Property.ApplicationPath + "KryptonParallelRecovery.exe";
                                    processInfo.CreateNoWindow = true;
                                    processInfo.WindowStyle = ProcessWindowStyle.Hidden;

                                    processInfo.Arguments = "\"" + Property.RecoverFromPopupFilepath + "/" + Property.ParallelRecoverySheetName + "\"" +
                                                      " " + "\"" + Common.Property.POPUP_SHEETNAME + "\"" +
                                                      " " + Process.GetCurrentProcess().Id.ToString();
                                    Common.Utility.setProcessParameter("KryptonParallelRecovery.exe");
                                    Process parallelProcess = Process.Start(processInfo);
                                    Thread.Sleep(1000);
                                    if (parallelProcess.Id == 0 || parallelProcess.HasExited)
                                    {
                                        Console.WriteLine(Common.exceptions.ERROR_AUTOIT);
                                        Console.WriteLine(Common.exceptions.ERROR_RECOVERY);
                                        return testSuiteResult;
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                //Please register AutoIt dll
                                Console.WriteLine(exceptions.ERROR_RECOVERY + e.Message);
                                Console.WriteLine(exceptions.ERROR_AUTOIT);
                                return testSuiteResult;
                            }
                            #endregion
                        }
                    }
                }
                catch (Exception exception)
                {
                    ReportHandler.WriteExceptionLog(new KryptonException(exception.Message), 0, exceptions.ERROR_OBJECTREPOSITORY);
                    return testSuiteResult;
                }
                #endregion
                if (testStepAction == null) //so that driver will not initialize every time
                    testStepAction = new TestDriver.Action(xmlRecoverFromPopupFlow, xmlRecoverFromBrowserFlow, ORTestData);
                try
                {
                    DateTime dtNow1 = DateTime.Now;
                    xmlLog.AddTestAttribute("ExecutionStartTime", dtNow1.ToString(Common.Property.DATE_TIME));
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    Console.WriteLine("Starting Test Case: " + testCases[testCaseCnt].Trim());
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    failureCount = 0;
                    stepNos = 0; //initialize the counter
                    Common.Property.EndExecutionFlag = false;

                    ExecuteTestCase(testCases[testCaseCnt].Trim()); //actual test flow

                    Common.Property.EndExecutionFlag = false;

                    DataSet xmlTestCaseDataSet = new DataSet(); //reset to null for the private global Test Case Data for iteration process 
                    //Shutting down the driver if CloseBrowserOnCompletion is set to true. 
                    if (Utility.GetParameter("closebrowseroncompletion").ToLower().Trim().Equals("true") || testCases.Length > 1)
                        testStepAction.Do("closeallbrowsers", "", "", "", "");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);

                    #region cleanup test flow
                    string[] cleanupCase = cleanupTestCase.Split(',');
                    string cleanupTestCaseId = string.Empty;
                    if (cleanupCase.Length > 0 && cleanupCase[0].ToLower().Contains("cleanuptestcaseid"))
                    {
                        cleanupTestCaseId = cleanupCase[0].Split('=')[1].Replace("\"", "");
                    }

                    if (cleanupCase.Length > 1 && cleanupCase[1].ToLower().Contains("data"))
                    {
                        string[] argument = cleanupTestCaseId.Split(Property.SEPERATOR);
                        for (int argCount = 0; argCount < argument.Length; argCount++)
                        {
                            Common.Utility.SetVariable("argument" + (argCount + 1), argument[argCount].Trim());
                        }
                    }

                    string cleanupTestCaseIteration = null;
                    if (cleanupCase.Length > 2 && cleanupCase[2].ToLower().Contains("iteration"))
                    {
                        cleanupTestCaseIteration = cleanupCase[2].Split('=')[1].Replace("\"", ""); ;
                    }
                    if (cleanupTestCaseId.IsNullOrWhiteSpace() == false)
                        ExecuteTestCase(cleanupTestCaseId, cleanupTestCaseIteration, testCaseId, "true"); 

                    #endregion


                    Console.WriteLine("Ending Test Case: " + testCases[testCaseCnt].Trim());
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    xmlLog.AddTestAttribute("RCMachineId", Common.Property.RCMachineId);

                    Common.Utility.ClearRunTimeDic(); //clear run time dictionary and copy parameter dictionary in run time dictionary

                    //Calculate test execution duration
                    TimeSpan tcDurationSpan = (DateTime.Now - dtNow1);
                    TimeSpan tcDuration = new TimeSpan(tcDurationSpan.Hours, tcDurationSpan.Minutes, tcDurationSpan.Seconds);
                    xmlLog.AddTestAttribute("ExecutionDuration", tcDuration.ToString());
                    Console.WriteLine("Total test execution duration was " + tcDuration.ToString());

                    dtNow1 = DateTime.Now;
                    xmlLog.AddTestAttribute("ExecutionEndTime", dtNow1.ToString(Common.Property.DATE_TIME));
                    xmlLog.AddTestAttribute("SauceJobUrl", Common.Property.JOB_URL);

                }
                catch (Exception exception)
                {
                    ReportHandler.WriteExceptionLog(new KryptonException(exception.Message), stepNos, "Execute Test Case");
                }
                Console.WriteLine(ConsoleMessages.MSG_DASHED);
                Console.WriteLine(ConsoleMessages.MSG_EXCECUTION_COMPLETE_LOG);
                Console.WriteLine(ConsoleMessages.MSG_DASHED);

                try
                {
                    //Generation of xml log file.
                    xmlLog.SaveXmlLog(); 

                    //for the last step all the file will upload after generation of html file.
                    if (testCaseCnt < testCases.Length - 1) 
                        Manager.UploadTestExecutionResults();
                }
                catch (Exception ex)
                {
                    logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
                }

                #region for Angieslist: Batch command to execution on test completion

                try
                {
                    string TestFinishCommand = Utility.GetVariable("TestCleanupBatch");
                    TestFinishCommand = Utility.ReplaceVariablesInString(TestFinishCommand);

                    if (!(TestFinishCommand.IsNullOrWhiteSpace()))
                    {
                        TestFinishCommand = Path.Combine(Common.Property.ApplicationPath, TestFinishCommand);
                        if (File.Exists(TestFinishCommand))
                        {
                            //Create process information and assign data
                            Process setupProcess = new Process();
                            setupProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            setupProcess.StartInfo.UseShellExecute = false;
                            setupProcess.StartInfo.FileName = TestFinishCommand;
                            setupProcess.StartInfo.WorkingDirectory = Common.Property.ApplicationPath;
                            setupProcess.StartInfo.ErrorDialog = false;

                            setupProcess.StartInfo.Arguments = Property.RCProcessId +
                                                               " " + "\"" + xmlLog.executedXmlPath + "\"";
                            // Start the process
                            setupProcess.Start();
                            //No need to wait for process to exit
                            setupProcess.WaitForExit(10000);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(exceptions.ERROR_CLEANUP + e.Message);
                }
                #endregion
            }

            //Stoping the IE Popup Thread after execution.
            executeThread = false;
            #region : Killing all the processes started by Krypton Application during execution.
            try
            {
                string[] procesesToKill = Common.Utility.getAllProcesses();
                for (int i = 0; i < procesesToKill.Length; i++)
                {
                    foreach (Process process in Process.GetProcessesByName(procesesToKill[i].Substring(0, procesesToKill[i].LastIndexOf('.'))))
                    {
                        try
                        {
                            process.Kill();
                            process.WaitForExit(10000);
                        }
                        catch (Exception ex)
                        { 
                            logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace); 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
            }
            #endregion

            //Shutting down the driver if CloseBrowserOnCompletion is set to true.
            if (Utility.GetParameter("Browser").Equals(Common.KryptonConstants.BROWSER_FIREFOX, StringComparison.OrdinalIgnoreCase))
            {
            }
            if (!Common.Utility.GetVariable("RCProcessId").IsNullOrWhiteSpace() && !Common.Utility.GetVariable("RCProcessId").Equals("RCProcessId", StringComparison.OrdinalIgnoreCase))
                testStepAction.Do("shutdowndriver"); //shutdown driver process running in remote

            TestDriver.Action.SaveScript();

            Console.WriteLine(ConsoleMessages.MSG_EXCECUTION_COMPLETED_HTML);

            //Execution end date and time set
            dtNow = DateTime.Now;
            Common.Property.ExecutionEndDateTime = dtNow.ToString(Common.Property.DATE_TIME);

            try
            {
               ReportHandler.CreateHtmlReportSteps(m_lstInputTestCaseIDs);
            }
            catch (Exception ex)
            {
                logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
                return testSuiteResult;
            }



            #region  Delete temp folder.
            /*List Of files that would be deleted here are :
            1. Text file for Error message storage that would be generated per test suit.(Size in KB).
            2. WebDriver Profile Folder that would be generated per test case.(Size ~ 23MB).
            */
            try
            {
                foreach (string file in Property.listOfFilesInTempFolder)
                {
                    if (Directory.Exists(file))
                    {
                        Directory.Delete(file, true);
                    }
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                }
            }
            catch (System.IO.IOException ex)
            {
                logwriter.WriteLog("IO Exception Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
            }
            catch (Exception ex)
            {
                Console.WriteLine(Utility.GetCommonMsgVariable("KRYPTONERRCODE0048"));
                logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
            }
            #endregion
            Process[] allProcess = Process.GetProcesses();
            foreach (Process process in allProcess)
            {
                if (process.ProcessName.Equals(KryptonConstants.CHROME_DRIVER, StringComparison.OrdinalIgnoreCase))
                {
                    process.Kill();
                }
                if (process.ProcessName.Equals(KryptonConstants.IE_DRIVER, StringComparison.OrdinalIgnoreCase))
                {
                    process.Kill();
                }
            }
            //Wait for user input at the end of the execution is handled by configuration file
            if (string.Equals(Common.Utility.GetParameter("EndExecutionWaitRequired"), "false", StringComparison.OrdinalIgnoreCase))
            {
                Array.ForEach(Process.GetProcessesByName("cmd"), x => x.Kill());
            }
            return testSuiteResult;
        }

        /// <summary>
        /// Set the projectpath and Set path for reading the .ini files 
        /// e.g Parameter.ini,EmailNotificaiton.ini etc
        /// </summary>
        /// <param string[]="args">Command line arguments contains the project properties.</param> 
        /// <returns></returns>
        private static void initializeIniPath(string[] args)
        {
            try
            {
                if (File.Exists(Path.Combine(Common.Property.ApplicationPath, "root.ini")))
                {
                    string currentProjectpath = string.Empty;
                    if (args.Length != 0)
                    {
                        for (int i = 0; i < args.Length; i++)
                        {
                            string[] KeyValuePair = args[i].Trim().Split('=');
                            string Key = KeyValuePair[0].Trim();
                            string Value = KeyValuePair[1].Trim();
                            if (Key.ToLower().Equals("projectpath") && !string.IsNullOrEmpty(Value))
                                currentProjectpath = Value;

                        }
                    }
                    if (currentProjectpath.Length == 0)
                    {
                        using (StreamReader sr = new StreamReader(Path.Combine(Common.Property.ApplicationPath, "root.ini")))
                        {
                            string currentProjectName = string.Empty;
                            while ((currentProjectName = sr.ReadLine()) != null)
                            {
                                if (currentProjectName.ToLower().Contains("projectpath"))
                                {
                                    currentProjectpath = currentProjectName.Substring(currentProjectName.IndexOf(':') + 1).Trim();
                                    break;
                                }
                            }
                        }
                        if (!Path.IsPathRooted(currentProjectpath))
                            currentProjectpath = Path.Combine(Common.Property.ApplicationPath, currentProjectpath);
                    }
                    Common.Property.IniPath = Path.GetFullPath(currentProjectpath);
                }
                else
                {
                    Common.Property.IniPath = Common.Property.ApplicationPath;
                }

            }
            catch (Exception e)
            {
                throw new Exception(exceptions.ERROR_INVALIDCOMMANDLINE);
            }
        }

        
        /// <summary>
        /// This is the main method which will fetch xml data from Test Manager and parse it
        /// and call action method of test driver, and set OR dictionary value for each step.
        /// After execution of each step it will call method of reporting and pass step execution result.
        /// Updated this method for CSV files.
        /// </summary>
        /// <param name="testCaseId">Test Case ID</param>
        /// <param name="iterationTestCase">Optional field, if non nagetive value is present, that means, it is for child 
        /// test case that have test data, it is like 1-1, 1-2, 2-4, ...</param>
        /// <param name="parentTestCaseId">Optional field, used for child test case data</param>
        /// <returns></returns>
        private static void ExecuteTestCase(string testCaseId, string iterationTestCase = null, string parentTestCaseId = null, string projectSpecific = null)
        {

            DataSet xmlTestCaseDataSet = new DataSet();
            xmlTestCaseDataSet = null; //initializing test data dataset
            string projectSpecificTestCaseFile = string.Empty; //initialization of project specific test case id
            int testCaseIdExists = -1;

            #region Retrieve test case workflow from test manager and set to local dataset
            DataSet xmlTestCaseFlow = TestCaseDataStruct();
            try
            {
                if (projectSpecific == "true")
                {
                    #region project specific reusable test cases fetching
                    string globalActionFileName = string.Empty;
                    if (!Path.GetExtension(Utility.GetParameter("reusabledefaultfile").ToLower().Trim()).Equals(Property.ExcelSheetExtension.ToLower()))
                        globalActionFileName = Utility.GetParameter("reusabledefaultfile") + Property.ExcelSheetExtension;
                    else
                        globalActionFileName = Utility.GetParameter("reusabledefaultfile").Trim();

                    string TempglobalActionFileName = globalActionFileName.Substring(0, globalActionFileName.LastIndexOf('.'));
                    if (Property.TestCaseIDParameter == 0)
                        projectSpecificTestCaseFile = Common.Utility.GetParameter("TestCaseId");
                    else
                        projectSpecificTestCaseFile = Property.TestCaseIDSeperator + TempglobalActionFileName;

                    xmlTestCaseFlow = objTestManager.GetTestCaseXml(Property.TestQuery, true, false);

                    //checking if test case id exists in the file or not
                    for (int drCnt = 0; drCnt < xmlTestCaseFlow.Tables[0].Rows.Count; drCnt++)
                    {
                        if (xmlTestCaseFlow.Tables[0].Rows[drCnt][KryptonConstants.TEST_CASE_ID].ToString().Equals(testCaseId, StringComparison.OrdinalIgnoreCase))
                        {
                            testCaseIdExists = drCnt;
                            break;
                        }
                    }
                    if (testCaseIdExists < 0)//if not exists in the main test flow
                    {
                        #region KRYPTON0253
                        string[] sep = { Property.TestCaseIDSeperator };
                        string[] testFlowSheetName = testCaseId.Split(sep, StringSplitOptions.None);
                        xmlTestCaseFlow = objTestManager.GetTestCaseXml(testFlowSheetName[0], false, true);
                        //checking if test case id exists in the file or not
                        for (int drCnt = 0; drCnt < xmlTestCaseFlow.Tables[0].Rows.Count; drCnt++)
                        {
                            if (xmlTestCaseFlow.Tables[0].Rows[drCnt][KryptonConstants.TEST_CASE_ID].ToString().Equals(testCaseId, StringComparison.OrdinalIgnoreCase))
                            {
                                testCaseIdExists = drCnt;
                                break;
                            }
                        }
                        if (testCaseIdExists < 0)
                        {
                            throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0049") + " [" + testCaseId + "] ", string.Empty);
                        }
                        #endregion
                    }
                    #endregion
                }
                else
                {
                    objTestManager = new Krypton.Manager(testCaseId);
                    xmlTestCaseFlow = objTestManager.GetTestCaseXml(Property.TestQuery, false, false);

                    //checking if test case id exists in the file or not
                    for (int drCnt = 0; drCnt < xmlTestCaseFlow.Tables[0].Rows.Count; drCnt++)
                    {
                        if (xmlTestCaseFlow.Tables[0].Rows[drCnt][KryptonConstants.TEST_CASE_ID].ToString().Equals(testCaseId, StringComparison.OrdinalIgnoreCase))
                        {
                            testCaseIdExists = drCnt;
                            break;
                        }
                    }
                    if (testCaseIdExists < 0)
                    {
                        throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0049") + " [" + testCaseId + "] ", string.Empty);
                    }
                }
            }
            catch (Exception exception)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0049") + " [" + testCaseId + "] ", exception.Message);
            }
            #endregion
            //Retrieve test case data if required from test manage based on iteration specified
            bool isTestDataUsed = false;
            int startingDataRow = 1;
            int endingDataRow = 1;

            if (iterationTestCase.IsNullOrWhiteSpace())
            {
                isTestDataUsed = false;
            }
            else
            {
                #region iteration step separtion
                //to fetch iteration step rows
                if (iterationTestCase.Contains('-'))
                    iterationTestCase = iterationTestCase.Replace("-", "^");

                string[] strIteration = iterationTestCase.Split('^');
                if (Regex.IsMatch(strIteration[0], "([0-9]*)"))
                {
                    startingDataRow = int.Parse(strIteration[0]);
                }
                if (strIteration.Length > 1)
                {
                    if (Regex.IsMatch(strIteration[1], "([0-9]*)"))
                    {
                        endingDataRow = int.Parse(strIteration[1]);
                    }
                }
                else
                {
                    endingDataRow = startingDataRow;
                }
                #endregion

                isTestDataUsed = true;
            }



            string xmlTestCaseName = null; //intialize of test case name
            string xmlTestCaseBrowser = null; //intialize of test case browser

            //Start a loop for all test data records
            for (int testDataCounter = startingDataRow; testDataCounter < endingDataRow + 1; testDataCounter++)
            {
                //Storing the Iteration number to a Dictionary key 'Iteration'. 
                Common.Utility.SetVariable("Iteration", testDataCounter.ToString());
                if (isTestDataUsed)
                {
                    xmlTestCaseDataSet = null;
                }

                bool currentTestCaseRecords = false; //to check if the loop reach the starting of the current test case records
                if (isTestDataUsed)
                {
                    xmlTestCaseDataSet = null;
                }
                //Start a loop for all test case flow record
                for (int caseRowCnt = testCaseIdExists; caseRowCnt < xmlTestCaseFlow.Tables[0].Rows.Count; caseRowCnt++)
                {
                    string xmlTestCaseId = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.TEST_CASE_ID].ToString();

                    #region to check record for current test case
                    if (xmlTestCaseId.IsNullOrWhiteSpace() == false && xmlTestCaseId.Equals(testCaseId, StringComparison.OrdinalIgnoreCase))
                        currentTestCaseRecords = true;

                    else if (xmlTestCaseId.IsNullOrWhiteSpace() == false && xmlTestCaseId.Equals(testCaseId, StringComparison.OrdinalIgnoreCase) == false)
                        currentTestCaseRecords = false;

                    if (xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.STEP_ACTION].ToString().Trim().IsNullOrWhiteSpace())
                        currentTestCaseRecords = false;

                    if (currentTestCaseRecords == false)
                        continue;
                    #endregion

                    //set test case name attribute to xml log file and that will be call once
                    if (xmlTestCaseName == null && parentTestCaseId.IsNullOrWhiteSpace())
                    {
                        xmlTestCaseName = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.TEST_SCENARIO].ToString();
                        xmlLog.AddTestAttribute("TestCase Name", xmlTestCaseName);
                    }

                    #region variable declaration and set value
                    string xmlTestCaseStepAction = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.STEP_ACTION].ToString().Trim();
                    string xmlTestCaseIteration = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.ITERATION].ToString().Trim();
                    string xmlTestCaseData = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.DATA].ToString().Trim();
                    string xmlTestCaseObj = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.TEST_OBJECT].ToString().Trim();
                    string xmlTestCaseParent = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.PARENT].ToString().Trim();
                    string xmlTestCaseOption = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.OPTIONS].ToString().Trim();
                    string xmlTestCaseComment = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt][KryptonConstants.COMMENTS].ToString().Trim();

                    //Replace variables in string here for all columns and pass on
                    xmlTestCaseStepAction = Common.Utility.ReplaceVariablesInString(xmlTestCaseStepAction);
                    xmlTestCaseData = Common.Utility.ReplaceVariablesInString(xmlTestCaseData);
                    xmlTestCaseParent = Common.Utility.ReplaceVariablesInString(xmlTestCaseParent);
                    xmlTestCaseObj = Common.Utility.ReplaceVariablesInString(xmlTestCaseObj);
                    xmlTestCaseOption = Common.Utility.ReplaceVariablesInString(xmlTestCaseOption);
                    xmlTestCaseComment = Common.Utility.ReplaceVariablesInString(xmlTestCaseComment);
                    xmlTestCaseIteration = Common.Utility.ReplaceVariablesInString(xmlTestCaseIteration);
                    #endregion

                    #region Read individual column values from Test Case Flow and pass on to test driver
                    if (xmlTestCaseStepAction.IsNullOrWhiteSpace() == false && ValidStepOption(xmlTestCaseOption) && Common.Property.EndExecutionFlag != true)
                    {
                        string cleanUpString = Common.Utility.GetCleanupTestCase(xmlTestCaseOption, "cleanuptestcaseid"); //capture Clean up test case id with options

                        if (!cleanUpString.Equals(string.Empty))
                            cleanupTestCase = cleanUpString;

                        string testCaseData = string.Empty;
                        if (xmlTestCaseStepAction.Equals("runExcelTestCase", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            //iteration method call for internal test case
                            ExecuteTestCase(xmlTestCaseData, xmlTestCaseIteration, testCaseId, null);
                        }
                        else
                        {
                            if (isTestDataUsed)
                            {

                                if (xmlTestCaseDataSet == null) // will call test data method only once
                                {
                                    #region call to test case data method to fetch corresponding data
                                    try
                                    {
                                        string testDataSheet = string.Empty;
                                        //fetching test data sheet name from the option field in the first line of the nested test case and set for the whole test case 
                                        string[] testDataSheet1 = Common.Utility.GetTestMode(xmlTestCaseOption,
                                                                                          "TestDataSheet").Split(',');
                                        if (testDataSheet1.Length > 1)
                                        {
                                            testDataSheet = testDataSheet1[1];
                                        }
                                        else
                                        {
                                            testDataSheet = "test_data";
                                        }
                                        Common.Property.TestDataSheet = testDataSheet;

                                        if (projectSpecific == "true")
                                        {
                                            xmlTestCaseDataSet = TestCaseData(parentTestCaseId,
                                                                        testDataCounter.ToString(),
                                                                        xmlTestCaseObj, true);
                                        }
                                        else
                                        {
                                            xmlTestCaseDataSet = TestCaseData(testCaseId, testDataCounter.ToString(),
                                                                        xmlTestCaseObj, false);
                                        }
                                    }
                                    catch (Exception exception)
                                    {
                                        stepNos++;
                                        throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0050"), exception.Message);
                                    }

                                    #endregion
                                }

                                if (xmlTestCaseFlow.Tables[0].Rows[caseRowCnt]["data"].ToString().IsNullOrWhiteSpace() == false)
                                {
                                    testCaseData = xmlTestCaseFlow.Tables[0].Rows[caseRowCnt]["data"].ToString();
                                }
                                else
                                {
                                    testCaseData = "{$[TD]" + xmlTestCaseObj + "}"; //to fetch by [td] variable
                                }
                            }

                            else
                                testCaseData = xmlTestCaseData;

                            #region Handling test data situations for verifyDatabase
                            if (xmlTestCaseStepAction.Equals("verifyDatabase", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                //for data from DBTestData sheet
                                try
                                {
                                    DBTestData(testCaseId, xmlTestCaseData, xmlTestCaseObj);
                                    testCaseData = xmlTestCaseData;
                                }
                                catch (Exception exception)
                                {
                                    stepNos++;
                                    throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0050"), exception.Message);
                                }
                            }
                            #endregion


                            testCaseData = testCaseData.Trim();

                            #region Starting special script based on data given in options column.

                            string[] scriptToRunContents = Common.Utility.GetTestMode(xmlTestCaseOption, "script").Split(',');
                            string scriptToRun = string.Empty;
                            if (scriptToRunContents.Length > 1)
                            {
                                scriptToRun = scriptToRunContents[1].ToString().ToLower();
                                try
                                {

                                    if (!(scriptToRun.Contains(".exe")))  // Updated to Check if script provided already contains an extension
                                    {
                                        scriptToRun = scriptToRun + ".exe";
                                    }
                                    Console.WriteLine("Executing special script: " + scriptToRun);

                                    Common.Utility.setProcessParameter(scriptToRun); //Add process name to process array before starting it.
                                    if ((scriptToRun.Equals("KillTPRecoveryforCSA.exe", StringComparison.OrdinalIgnoreCase)) || (scriptToRun.Equals("KillParallelRecovery.exe", StringComparison.OrdinalIgnoreCase)))
                                    {
                                        foreach (Process process in Process.GetProcesses())
                                        {
                                            if (process.ProcessName.Equals("KryptonParallelRecovery", StringComparison.OrdinalIgnoreCase))
                                            {
                                                process.Kill();
                                                break;
                                            }

                                        }
                                    }
                                    else
                                    {
                                        // Create process information and assign data
                                        Process specialScriptProcess = new Process();
                                        specialScriptProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                                        specialScriptProcess.StartInfo.UseShellExecute = false;
                                        specialScriptProcess.StartInfo.FileName = scriptToRun;
                                        specialScriptProcess.StartInfo.ErrorDialog = false;

                                        // Start the process
                                        specialScriptProcess.Start();
                                        specialScriptProcess.WaitForExit(10000);
                                    }
                                }
                                catch (Exception) { }
                            }
                            #endregion
                            //Replace variables in string here for all columns and pass on
                            xmlTestCaseStepAction = Common.Utility.ReplaceVariablesInString(xmlTestCaseStepAction);
                            xmlTestCaseParent = Common.Utility.ReplaceVariablesInString(xmlTestCaseParent);
                            xmlTestCaseObj = Common.Utility.ReplaceVariablesInString(xmlTestCaseObj);
                            xmlTestCaseOption = Common.Utility.ReplaceVariablesInString(xmlTestCaseOption);
                            xmlTestCaseComment = Common.Utility.ReplaceVariablesInString(xmlTestCaseComment);
                            xmlTestCaseIteration = Common.Utility.ReplaceVariablesInString(xmlTestCaseIteration);
                            testCaseData = Common.Utility.ReplaceVariablesInString(testCaseData);

                            if (xmlTestCaseOption.IndexOf("{append}", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                testCaseData = Common.Utility.ReplaceVariablesInString("{$" + xmlTestCaseObj + "}") + testCaseData;
                            }


                            //Replace variable with special charecters)

                            if (ValidStepOption(testCaseData) == false) continue;

                            #region call to TestDriver Action.Do method
                            try
                            {

                                //Initialize Step Log.
                                Common.Property.InitializeStepLog();
                                if (xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0 && !(xmlTestCaseStepAction.ToLower().Equals("setvariable") && Common.Property.DebugMode.Equals("false")))
                                {
                                    stepNos++;
                                    Common.Property.StepNumber = stepNos.ToString();
                                }
                                Console.WriteLine("Parent:" + xmlTestCaseParent);
                                Console.WriteLine("Object:" + xmlTestCaseObj);
                                Console.WriteLine("Step Action:" + xmlTestCaseStepAction);
                                if (testCaseData.IndexOf("{$[TD]", StringComparison.OrdinalIgnoreCase) < 0)
                                    Console.WriteLine("Data:" + testCaseData);
                                else Console.WriteLine("Data:");

                                //Set TestDriver object dictionary
                                TestDriver.Action.objDataRow = GetTestOrData(xmlTestCaseParent, xmlTestCaseObj);

                                // In case of DragAndDrop, there are two objects to be handled simultaneously.
                                // Set second object dictionary 
                                if (xmlTestCaseStepAction.Equals("DragAndDrop", StringComparison.OrdinalIgnoreCase) == true)
                                    TestDriver.Action.objSecondDataRow = GetTestOrData(xmlTestCaseParent, xmlTestCaseData);

                                if (TestDriver.Action.objDataRow.ContainsKey(KryptonConstants.WHAT))
                                    Console.WriteLine("OR Locator:" + TestDriver.Action.objDataRow[KryptonConstants.WHAT]);
                                else
                                    Console.WriteLine("OR Locator:");

                                Common.Property.ExecutionDate = DateTime.Now.ToString(Common.Utility.GetParameter("DateFormat"));
                                Common.Property.ExecutionTime = DateTime.Now.ToString(Common.Utility.GetParameter("TimeFormat"));

                                testStepAction.Do(xmlTestCaseStepAction, xmlTestCaseParent, xmlTestCaseObj,
                                                   testCaseData,
                                                   xmlTestCaseOption);

                                //if action method is not implemented is returned, need to check for project specific functions
                                if (Common.Property.Remarks.IndexOf(Utility.GetCommonMsgVariable("KRYPTONERRCODE0025"), StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    #region checkprojectspecificfunction
                                    if (xmlTestCaseStepAction.IndexOf(Common.Property.TestCaseIDSeperator,
                                                                      StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        if (xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0)
                                        {
                                            stepNos--;
                                        }

                                        #region KRYPTON0252

                                        string[] argument = testCaseData.Split(Property.SEPERATOR);
                                        if (argument.Length > 0)
                                        {

                                            for (int argCount = 0; argCount < argument.Length; argCount++)
                                            {
                                                Common.Utility.SetVariable("argument" + (argCount + 1), argument[argCount].Trim());
                                            }

                                        }
                                        #endregion

                                        ExecuteTestCase(xmlTestCaseStepAction, xmlTestCaseIteration, testCaseId,
                                                        "true");
                                        continue;
                                    }

                                    #endregion
                                }
                                {
                                    Boolean flag = false;
                                    if (xmlTestCaseOption.Equals("{NoReporting}") && Common.Property.Status == Common.ExecutionStatus.Fail)
                                        flag = true;
                                    if (xmlTestCaseOption.Equals("{NoReporting}").Equals(false))
                                        flag = true;
                                    if ((xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0))
                                    {
                                        if (flag)
                                        {
                                            if (Common.Property.Status == Common.ExecutionStatus.Fail &&
                                                xmlTestCaseStepAction.IndexOf("verify",
                                                                              StringComparison.OrdinalIgnoreCase) < 0)
                                            {
                                                testSuiteResult = 1; //process fail
                                                //Increment Failure Count only in case of non optional and non ignored step cases.

                                                failureCount++;

                                                if ((failureCount ==
                                                    int.Parse(Common.Utility.GetParameter("FailedCountForExit")) &&
                                                    int.Parse(Common.Utility.GetParameter("FailedCountForExit")) > 0)
                                                    || Common.Property.EndExecutionFlag == true)
                                                {
                                                    if (Common.Property.EndExecutionFlag == true)
                                                        Console.WriteLine(ConsoleMessages.MSG_EXCEUTION_ENDS);
                                                    else
                                                        Console.WriteLine(ConsoleMessages.MSG_FAILURE_EXCEED);

                                                    Common.Property.StepComments = xmlTestCaseComment +
                                                        ": End execution due to nos. of failure(s) exceeds failed counter defined in parameter file -" + failureCount + ".";
                                                    xmlLog.WriteExecutionLog();

                                                    Common.Property.EndExecutionFlag = true;
                                                    return;
                                                }
                                                else
                                                {
                                                    Common.Property.StepComments = xmlTestCaseComment;
                                                    xmlLog.WriteExecutionLog();
                                                    Console.WriteLine("Status:" + Common.Property.Status);
                                                    Console.WriteLine("Remarks:" + Common.Property.Remarks);
                                                }

                                            }
                                            else
                                            {
                                                Common.Property.StepComments = xmlTestCaseComment;
                                                if (!(xmlTestCaseStepAction.ToLower().Equals("setvariable") && Common.Property.DebugMode.Equals("false")))
                                                    xmlLog.WriteExecutionLog();
                                                Console.WriteLine("Status:" + Common.Property.Status);
                                                Console.WriteLine("Remarks:" + Common.Property.Remarks);
                                            }
                                        }
                                    }
                                    //set browser attribute to xml log file and that will be call once
                                    if (xmlTestCaseBrowser == null && !Common.Utility.GetVariable(Property.BrowserVersion).Equals("version", StringComparison.OrdinalIgnoreCase))
                                    {
                                        xmlTestCaseBrowser = Common.Utility.GetVariable(Common.Property.BrowserString) + Common.Utility.GetVariable(Property.BrowserVersion);
                                        xmlLog.AddTestAttribute("Browser", Common.Utility.GetVariable(Common.Property.BrowserString) + Common.Utility.GetVariable(Property.BrowserVersion));
                                    }
                                }
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {
                                throw new Exception("Syntax error in test input.");
                            }
                            catch (Exception exception)
                            {
                                throw new Exception(Utility.GetCommonMsgVariable("KRYPTONERRCODE0051").Replace("{MSG}", exception.Message));
                            }

                            Console.WriteLine(ConsoleMessages.MSG_DASHED);

                            if (xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                xmlLog.SaveXmlLog(); //Generation of xml log file 
                            }

                            #endregion
                        }
                    }
                }
            }
        }


        #region Test Case Data structure and Data
        /// <summary>
        /// This method will fetch test data of the iteration test case
        /// </summary>
        /// <param name="childTestCaseId">Test Case Id for iteration</param>
        /// <param name="rowNo">Get data from which row, it is defined by the iteration 1-1, 2-2, ...</param>
        /// <param name="testObject">Get data for which object</param>
        /// <returns>void</returns>
        private static DataSet TestCaseData(string childTestCaseId, string rowNo, string testObject, bool projectspecific)
        {
            DataSet xmlTestCaseDataSet = new DataSet();
            try
            {
                {
                    string globalActionFileName = string.Empty;
                    if (!Path.GetExtension(Utility.GetParameter("reusabledefaultfile").ToLower().Trim()).Equals(Property.ExcelSheetExtension.ToLower()))
                        globalActionFileName = Utility.GetParameter("reusabledefaultfile") + Property.ExcelSheetExtension;
                    else
                        globalActionFileName = Utility.GetParameter("reusabledefaultfile").Trim();

                    Krypton.Manager objTestManager = null;
                    if (projectspecific)
                    {
                        objTestManager = new Krypton.Manager(childTestCaseId, globalActionFileName);
                    }
                    else
                    {
                        objTestManager = new Krypton.Manager(childTestCaseId);
                    }
                    xmlTestCaseDataSet = objTestManager.GetTestDataXml();
                    //Set test data in the runtime dictionary for parameter calls in the test case by {$[TD]##$}

                    for (int dsDataCnt = 0; dsDataCnt < xmlTestCaseDataSet.Tables[0].Rows.Count; dsDataCnt++)
                    {
                        if (xmlTestCaseDataSet.Tables[0].Rows[dsDataCnt]["row_No"].ToString().Equals(rowNo))
                        {
                            for (int dsColumnCnt = 0;
                                 dsColumnCnt < xmlTestCaseDataSet.Tables[0].Columns.Count;
                                 dsColumnCnt++)
                            {
                                string columnName = "[TD]" +
                                                    xmlTestCaseDataSet.Tables[0].Columns[dsColumnCnt].ColumnName;
                                //get Column name
                                string dataValue = xmlTestCaseDataSet.Tables[0].Rows[dsDataCnt][dsColumnCnt].ToString();
                                //get data
                                dataValue = Utility.ReplaceVariablesInString(dataValue);
                                Common.Utility.SetVariable(columnName, dataValue); //set value
                            }
                        }
                    }

                }
                //
            }
            catch (Common.KryptonException exception)
            {
                throw exception;
            }

            return xmlTestCaseDataSet;


        }

        /// <summary>
        /// This method will create the structure for Test Case DataSet
        /// </summary>
        /// <returns>Null DataSet</returns>
        private static DataSet TestCaseDataStruct()
        {

            DataSet xmlTestCaseFlow = new DataSet("TestCaseFlow");
            DataTable dtTestCase = new DataTable("TestCase");
            DataColumn dcKeyword = new DataColumn(KryptonConstants.KEYWORDS);
            DataColumn dcTestScenario = new DataColumn(KryptonConstants.TEST_SCENARIO);
            DataColumn dcTestCaseId = new DataColumn(KryptonConstants.TEST_CASE_ID);
            DataColumn dcComments = new DataColumn(KryptonConstants.COMMENTS);
            DataColumn dcParent = new DataColumn(KryptonConstants.PARENT);
            DataColumn dcTestObject = new DataColumn(KryptonConstants.TEST_OBJECT);
            DataColumn dcsTepAction = new DataColumn(KryptonConstants.STEP_ACTION);
            DataColumn dcData = new DataColumn(KryptonConstants.DATA);
            DataColumn dcIteration = new DataColumn(KryptonConstants.ITERATION);
            DataColumn dcOptions = new DataColumn(KryptonConstants.OPTIONS);

            dtTestCase.Columns.Add(dcKeyword);
            dtTestCase.Columns.Add(dcTestScenario);
            dtTestCase.Columns.Add(dcTestCaseId);
            dtTestCase.Columns.Add(dcComments);
            dtTestCase.Columns.Add(dcParent);
            dtTestCase.Columns.Add(dcTestObject);
            dtTestCase.Columns.Add(dcsTepAction);
            dtTestCase.Columns.Add(dcData);
            dtTestCase.Columns.Add(dcIteration);
            dtTestCase.Columns.Add(dcOptions);

            xmlTestCaseFlow.Tables.Add(dtTestCase);

            return xmlTestCaseFlow;

        }
        #endregion

        #region Database Test Data structure and Data
        /// <summary>
        /// This method will fetch test data of the iteration test case
        /// </summary>
        /// <param name="testCaseId">Test Case Id for iteration</param>
        /// <param name="rowNo">Get data from which row, it is defined by the iteration 1-1, 2-2, ...</param>
        /// <param name="testObject">Get data for which object</param>
        /// <returns>string data</returns>
        private static void DBTestData(string testCaseId, string rowNo, string testObject)
        {
            try
            {
                Krypton.Manager objTestManager = new Krypton.Manager(testCaseId);
                xmlTestDataSet = objTestManager.GetDBTestDataXml();
                //Set test data in the runtime dictionary for parameter calls in the test case by {$[TD]##$}
                for (int dsDataCnt = 0; dsDataCnt < xmlTestDataSet.Tables[0].Rows.Count; dsDataCnt++)
                {
                    if (xmlTestDataSet.Tables[0].Rows[dsDataCnt]["SNo"].ToString().Equals(rowNo))
                    {
                        for (int dsColumnCnt = 0;
                             dsColumnCnt < xmlTestDataSet.Tables[0].Columns.Count;
                             dsColumnCnt++)
                        {
                            string columnName = "[TD]" +
                                                xmlTestDataSet.Tables[0].Columns[dsColumnCnt].ColumnName;
                            //get Column name
                            string dataValue = xmlTestDataSet.Tables[0].Rows[dsDataCnt][dsColumnCnt].ToString();
                            //get data
                            Common.Utility.SetVariable(columnName, dataValue);
                        }
                    }
                }
            }
            catch (Common.KryptonException exception)
            {
                throw exception;
            }
        }
        #endregion
                    #endregion



        #region OR data
        /// <summary>
        /// This method will create the dictionary of OR for the passing parent and object 
        /// </summary>
        /// <param name=KRYPTONConstants.PARENT>Filter parent in the OR dataset</param>
        /// <param name="testObj">Filter test_objecy in the OR dataset</param>
        /// <returns>Dictionary with key=>value pairs
        /// Keys are=>logical_name,obj_type,how,what,mapping
        /// </returns>
        private static Dictionary<string, string> GetTestOrData(string parent, string testObj)
        {
            Dictionary<string, string> orDataRow = new Dictionary<string, string>();
            //Assigning testObj to parent value if testObj is not given.
            if (testObj.Equals(String.Empty))
                testObj = parent;
            if (orDataRow.ContainsKey(KryptonConstants.LOGICAL_NAME))
            {
                orDataRow[KryptonConstants.LOGICAL_NAME] = string.Empty;
                orDataRow[KryptonConstants.OBJ_TYPE] = string.Empty;
                orDataRow[KryptonConstants.HOW] = string.Empty;
                orDataRow[KryptonConstants.WHAT] = string.Empty;
                orDataRow[KryptonConstants.MAPPING] = string.Empty;
                orDataRow[KryptonConstants.TEST_OBJECT] = string.Empty;
            }
            try
            {

                foreach (DataRow drData in ORTestData.Tables[0].Rows)
                {

                    if (drData[KryptonConstants.PARENT].ToString().Trim().Equals(parent, StringComparison.OrdinalIgnoreCase)
                        && drData[KryptonConstants.TEST_OBJECT].ToString().Trim().Equals(testObj, StringComparison.OrdinalIgnoreCase)
                        && drData[KryptonConstants.TEST_OBJECT].ToString().Trim().IsNullOrWhiteSpace() == false)
                    {
                        if (orDataRow.ContainsKey(KryptonConstants.LOGICAL_NAME) == false)
                        {
                            orDataRow.Add(KryptonConstants.PARENT, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.PARENT].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.LOGICAL_NAME, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.LOGICAL_NAME].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.OBJ_TYPE, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.HOW, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.WHAT, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.MAPPING, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.TEST_OBJECT, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.TEST_OBJECT].ToString().Trim()));//  to retrieve object in the HTML also.(PBI 118 in TFs)
                        }
                        else
                        {
                            orDataRow[KryptonConstants.PARENT] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.PARENT].ToString().Trim());
                            orDataRow[KryptonConstants.LOGICAL_NAME] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.LOGICAL_NAME].ToString().Trim());
                            orDataRow[KryptonConstants.OBJ_TYPE] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString().Trim());
                            orDataRow[KryptonConstants.HOW] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString().Trim());
                            orDataRow[KryptonConstants.WHAT] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString().Trim());
                            orDataRow[KryptonConstants.MAPPING] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString().Trim());
                            orDataRow[KryptonConstants.TEST_OBJECT] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.TEST_OBJECT].ToString().Trim());// to retrieve object in the HTML also.(PBI 118 in TFs)
                        }
                    }

                }
            }
            catch (Common.KryptonException exception)
            {
                throw new TargetException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0052").Replace("{MSG1}", parent).Replace("{MSG2}", testObj));
            }


            return orDataRow;
        }

        /// <summary>
        /// This method will fetch Object Repository data from manager class
        /// and set it as private static variable, so that it can be accessed from other private methods
        /// </summary> 
        /// <returns>Null DataSet</returns>
        private static void OrData()
        {
            try
            {
                DataSet dsOR = OrDataStruct();
                dsOR = Manager.GetObjectDefinitionXml();
                ORTestData = dsOR;
                Property.orBachupData = dsOR;
            }
            catch (Common.KryptonException exception)
            {

                throw exception;
            }
        }

        /// <summary>
        /// This method will create the structure for Object Repository DataSet
        /// </summary>
        /// <returns>Null DataSet</returns>
        private static DataSet OrDataStruct()
        {

            DataSet dsOrStruct = new DataSet("ORDataSet");
            DataTable dtOrTable = new DataTable("ORTable");
            DataColumn dcSlNo = new DataColumn("sl_no");
            DataColumn dcParent = new DataColumn(KryptonConstants.PARENT);
            DataColumn dtTestObj = new DataColumn(KryptonConstants.TEST_OBJECT);
            DataColumn dcLogicName = new DataColumn(KryptonConstants.LOGICAL_NAME);
            DataColumn dcLocale = new DataColumn(KryptonConstants.LOCALE);
            DataColumn dcObjType = new DataColumn(KryptonConstants.OBJ_TYPE);
            DataColumn dcHow = new DataColumn(KryptonConstants.HOW);
            DataColumn dcWhat = new DataColumn(KryptonConstants.WHAT);
            DataColumn dcComment = new DataColumn(KryptonConstants.COMMENTS);
            DataColumn dcMapping = new DataColumn(KryptonConstants.MAPPING);

            dtOrTable.Columns.Add(dcSlNo);
            dtOrTable.Columns.Add(dcParent);
            dtOrTable.Columns.Add(dtTestObj);
            dtOrTable.Columns.Add(dcLogicName);
            dtOrTable.Columns.Add(dcLocale);
            dtOrTable.Columns.Add(dcObjType);
            dtOrTable.Columns.Add(dcHow);
            dtOrTable.Columns.Add(dcWhat);
            dtOrTable.Columns.Add(dcComment);
            dtOrTable.Columns.Add(dcMapping);
            dsOrStruct.Tables.Add(dtOrTable);

            return dsOrStruct;

        }

        #endregion


        /// <summary>
        /// To have a generalised step validation
        /// </summary>
        /// <param name="stepOption"></param>
        /// <returns>true/false</returns>
        private static bool ValidStepOption(string stepOption)
        {
            bool validStep = true;

            if ((stepOption.IndexOf("{ignore}", StringComparison.OrdinalIgnoreCase) >= 0 && stepOption.IndexOf("{ignorespace}", StringComparison.OrdinalIgnoreCase) < 0 && stepOption.IndexOf("{ignorecase}", StringComparison.OrdinalIgnoreCase) < 0) ||
                stepOption.IndexOf("{skip}", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                validStep = false;
            }

            string[] testmode = Common.Utility.GetTestMode(stepOption, Common.Property.TESTMODE).Split(',');

            if (testmode.Length > 1)
            {
                if (testmode[1].Equals(Utility.GetVariable(testmode[0]), StringComparison.OrdinalIgnoreCase) == false)
                {
                    validStep = false;
                }
            }
            return validStep;
        }

    }

}