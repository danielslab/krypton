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
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Common;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Configuration;
using System.Threading;
using Microsoft.Practices.Unity;

namespace Krypton
{
    class TestEngine : Variables
    {

        #region Member Variables
        private static DataSet _orTestData = null;
        private static int _stepNos = 0;
        private static int _failureCount = 0;
        private static DataSet _xmlTestDataSet = null;
        private static Manager _objTestManager = null;
        private static string _cleanupTestCase = string.Empty;
        public static IKryptonLogger Logwriter = null;
        public static bool ExecuteThread = false;
        #endregion

        public static IUnityContainer container;
        static void RegisterDependencies()
        {
            container = new UnityContainer();
            container.RegisterType<Manager>();
            container.RegisterType<IKryptonLogger, KryptonFileLogWriter>();
        }


        /// <summary>
        /// This is Main console engine
        /// execute test cases and generate log file
        /// </summary>
        /// <param name="args">console arguments which will override parameter.ini default parameters</param> 
        /// <returns></returns>
        static int Main(string[] args)
        {
            string version = ConfigurationSettings.AppSettings["Version"];
            RegisterDependencies();
            Console.Title = "Krypton" + " v" + "2.0.1";
            try
            {
                //To find application startup path.
                string applicationPath = Application.StartupPath;
                if (applicationPath.Length > 0)
                {
                    Property.ApplicationPath = applicationPath + "\\";
                }
                else
                {
                    DirectoryInfo dr = new DirectoryInfo("./");
                    Property.ApplicationPath = dr.FullName;
                }

                InitializeIniPath(args);
                Utility.SetParameter("ApplicationPath", Property.ApplicationPath);
                Utility.SetVariable("ApplicationPath", Property.ApplicationPath);

                Utility.GetCommonMessageData();

                //Parse ini file and put all parameters into global dictionary

                #region Adding internal parameter, runbyevents.
                /* 
                   Parameter "runbyevents" means, certain methods will run using browser events instead of using native driver implemented methods.
                   This can also be passed from command line arguments, but is not added in parameter.ini file.
                   When its value is true, browser events will be used to perform operations instead of native methods.
                   E.g. arguments[0].click(); will be used instead of testobject.click when this parameter is true
                */
                Property.Parameterdic.Add("runbyevents", "false");
                Property.Runtimedic.Add("runbyevents", "false");
                Utility.SetParameter("keyword", string.Empty);
                Utility.SetVariable("keyword", string.Empty);
                Utility.SetParameter("currentRelease", "0");
                Utility.SetVariable("currentRelease", "0");
                if (Path.IsPathRooted(Property.DestinationFolderDownload))
                {
                    Utility.SetParameter("downloadpath", Property.DestinationFolderDownload);
                    Utility.SetVariable("downloadpath", Property.DestinationFolderDownload);
                }
                else
                {
                    Utility.SetParameter("downloadpath", string.Concat(Property.ApplicationPath, Property.DestinationFolderDownload));
                    Utility.SetVariable("downloadpath", string.Concat(Property.ApplicationPath, Property.DestinationFolderDownload));
                }
                #endregion

                Utility.CollectkeyValuePairs();

                #region Parse arguments passed to executable. This may over-write ini file parameters, if names are same
                if (args.Length != 0)
                {
                    for (int i = 0; i < args.Length; i++)
                    {
                        string[] keyValuePair = args[i].Trim().Split('=');
                        string key = keyValuePair[0].Trim().ToLower();
                        string value = keyValuePair[1].Trim();
                        Utility.SetParameter(key, value);
                        Utility.SetVariable(key, value);
                    }
                }
                #endregion
                // Set Environment File Paths
                Property.EnvironmentFileLocation = Path.IsPathRooted(Utility.GetParameter("EnvironmentFileLocation")) ? Utility.GetParameter("EnvironmentFileLocation") : Path.Combine(Property.IniPath, Utility.GetParameter("EnvironmentFileLocation"));

                Utility.UpdateKeyValuePairs();
                Utility.ValidateParameters();
                Property.RecoveryCount = Convert.ToInt16(Utility.GetParameter("recoverycount"));
                // modified to set the value only if it is given in parameters.ini
                if (Utility.GetParameter("globaltimeout") != string.Empty)
                    Property.GlobalTimeOut = Utility.GetParameter("globaltimeout");


                //In case manager type is MSTestManager, user can pass on any of the following options:
                //MSTestManager,MSTM,MTM,VSTM,VSTestManager
                if (Utility.GetParameter("ManagerType").StartsWith("MSTestManager", StringComparison.OrdinalIgnoreCase) ||
                    Utility.GetParameter("ManagerType").StartsWith("MSTM", StringComparison.OrdinalIgnoreCase) ||
                    Utility.GetParameter("ManagerType").StartsWith("MTM", StringComparison.OrdinalIgnoreCase) ||
                    Utility.GetParameter("ManagerType").StartsWith("VSTM", StringComparison.OrdinalIgnoreCase) ||
                    Utility.GetParameter("ManagerType").StartsWith("VSTestManager", StringComparison.OrdinalIgnoreCase))
                {
                    Utility.SetParameter("ManagerType", "MSTestManager");
                    Utility.SetVariable("ManagerType", "MSTestManager");
                }
                //Get browser settings
                Utility.InitializeBrowser(Utility.GetParameter(Property.BrowserString));

                Property.ApplicationUrl = Utility.GetParameter("ApplicationURL");

                int num;
                if (Property.Parameterdic.ContainsKey("waitforalert"))
                    if (!string.IsNullOrEmpty(Utility.GetParameter("Waitforalert")) && int.TryParse(Utility.GetParameter("Waitforalert"), out num))
                        Property.Waitforalert = int.Parse(Utility.GetParameter("Waitforalert"));

                Property.ExcelSheetExtension = Utility.GetParameter("TestCaseFileExtension");
                if (!Property.ExcelSheetExtension.StartsWith("."))
                {
                    Property.ExcelSheetExtension = "." + Property.ExcelSheetExtension;
                }

                //Get Test Case File path
                Property.TestCaseFilepath = Path.IsPathRooted(Utility.GetParameter("TestCaseLocation")) ? Utility.GetParameter("TestCaseLocation") : Path.Combine(Property.IniPath, Utility.GetParameter("TestCaseLocation"));

                //set the ManagerType
                if (Utility.GetParameter("ManagerType").IsNullOrWhiteSpace() == false)
                    Property.ManagerType = Utility.GetParameter("ManagerType");

                if (Utility.GetParameter("ListOfUniqueCharacters").IsNullOrWhiteSpace() == false)
                {
                    Property.ListOfUniqueCharacters = Utility.GetParameter("ListOfUniqueCharacters").Split(',');
                }

                //Handle Test Case Id conditions
                #region check for test suite
                if (Utility.GetParameter("TestSuite").Trim().IsNullOrWhiteSpace() && Utility.GetParameter("TestSuite") != string.Empty)
                {
                    Utility.SetParameter("TestSuite", Utility.GetParameter("TestCaseId").Trim().Split('.')[0]);
                }
                else
                {
                    TestSuite = Utility.GetParameter("TestSuite").Trim();
                }
                if (Property.Parameterdic["testsuite"] != string.Empty && Property.Parameterdic["testcaseid"] != string.Empty)
                {
                    Property.Runtimedic["testcaseid"] = Property.Parameterdic["testcaseid"];
                    Utility.SetParameter("TestSuite", string.Empty);
                    Property.Runtimedic["testsuite"] = string.Empty;
                }
                if (Utility.GetParameter("TestSuite") != string.Empty)
                {
                    var testsuiteFilePath =
                        Path.GetFullPath(Property.IniPath + "/" + Utility.GetParameter("TestSuiteLocation") + "/" +
                                         Utility.GetParameter("TestSuite"));
                    //Read the suite excel file to dataset
                    DataSet suiteFileDataSet = Manager.GetSuiteDataSet(testsuiteFilePath);
                    Utility.SetParameter("TestCaseId", Utility.GetTestCaseIdFromSuiteDataset(suiteFileDataSet));
                    Property.Runtimedic["testcaseid"] = string.Empty;
                }
                if (!Utility.GetParameter("TestCaseId").Trim().IsNullOrWhiteSpace())
                    TestSuite = Utility.GetParameter("TestCaseId").Trim().Split('.')[0];
                TestCaseId = Utility.GetParameter("TestCaseId");
                Manager.IsTestSuite = false;

                //Test Cases that Input From the Parameter.ini
                #endregion
                //Test Cases that are exit in file
                string[] matchingTestCase = TestCaseId.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                string finalTestCases = null;
                finalTestCases = Property.Runtimedic["testcaseid"];
                if (Utility.GetParameter("TestSuite") != string.Empty)
                {
                    finalTestCases = matchingTestCase.Aggregate(finalTestCases, (current, item) => current + "," + item);
                    finalTestCases = finalTestCases.Substring(1);
                }

                TestCases = finalTestCases.Split(',');
                // Check if there is at least one test case id for test execution
                if (TestCases.Length.Equals(0))
                {
                    Console.WriteLine(Exceptions.ERROR_NOTESTCASEFOUND);
                    return TestSuiteResult;
                }

                KryptonUtility.SetProjectFilesPaths();


                #region to get database from different ini files
                string[] sqlConnectionStr = Property.DbConnectionString.Trim().Split(';');

                for (int sqlParamCnt = 0; sqlParamCnt < sqlConnectionStr.Length; sqlParamCnt++)
                {
                    if (sqlConnectionStr[sqlParamCnt].Trim().IndexOf("database", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string oriDatabase = sqlConnectionStr[sqlParamCnt];
                        sqlConnectionStr[sqlParamCnt] = Utility.GetParameter("write").IsNullOrWhiteSpace() == false
                            ? "database=" + Utility.GetParameter("write")
                            : oriDatabase;

                        Property.SqlConnectionStringWrite = string.Join(";", sqlConnectionStr);
                        sqlConnectionStr[sqlParamCnt] = Utility.GetParameter("read").IsNullOrWhiteSpace() == false
                            ? "database=" + Utility.GetParameter("read")
                            : oriDatabase;
                        Property.SqlConnectionStringRead = string.Join(";", sqlConnectionStr);
                    }
                }
                #endregion
                //dbtestdata file path
                Property.SqlQueryFilePath = Path.IsPathRooted(Utility.GetParameter("DBQueryFilePath")) ? Utility.GetParameter("DBQueryFilePath") : string.Concat(Property.IniPath, Utility.GetParameter("DBQueryFilePath"));

                //to split testcaseid for test case excel file
                Property.TestCaseIdSeperator = Utility.GetParameter("TestCaseIDSeperator");

                //to fetch test case excel file from testcaseid
                Property.TestCaseIdParameter = int.Parse(Utility.GetParameter("TestCaseIDParameter"));

                Property.DebugMode = Utility.GetParameter("debugmode").Trim();

                //for snapshot option
                Property.SnapshotOption = Utility.GetParameter("SnapshotOption");

                //for remote execution initialization
                Property.IsRemoteExecution = Utility.GetParameter("RunRemoteExecution");

                //Force snapshot option in case of remote driver
                if (Property.IsRemoteExecution.Equals("true") && Property.SnapshotOption.Equals("always"))
                {
                    Property.SnapshotOption = "on page change";
                }

                Property.RemoteUrl = Utility.GetParameter("RunOnRemoteBrowserUrl");

                //for email templates
                Property.EmailStartTemplate = Path.IsPathRooted(Utility.GetParameter("EmailStartTemplate")) ? Utility.GetParameter("EmailStartTemplate") : string.Concat(Property.IniPath, Utility.GetParameter("EmailStartTemplate"));

                Property.EmailEndTemplate = Path.IsPathRooted(Utility.GetParameter("EmailEndTemplate")) ? Utility.GetParameter("EmailEndTemplate") : string.Concat(Property.IniPath, Utility.GetParameter("EmailEndTemplate"));


                //scripting language
                if (!Utility.GetVariable("ScriptLanguage").IsNullOrWhiteSpace())
                    Property.ScriptLanguage = Utility.GetVariable("ScriptLanguage");
            }
            catch (Exception exception)
            {
                Console.WriteLine(new KryptonException(exception.Message));
                return TestSuiteResult;
            }
            if (!Utility.GetVariable("RCProcessId").IsNullOrWhiteSpace() && !Utility.GetVariable("RCProcessId").Equals("RCProcessId", StringComparison.OrdinalIgnoreCase))
                Property.RcProcessId = Utility.GetVariable("RCProcessId");

            if (!Utility.GetVariable("RCMachineId").IsNullOrWhiteSpace() && !Utility.GetVariable("RCMachineId").Equals("RCMachineId", StringComparison.OrdinalIgnoreCase))
            {
                Property.RcMachineId = Utility.GetVariable("RCMachineId");

                //Update saucelabs machine status to protect identity of username and api key
                if (Property.RcMachineId.ToLower().Contains(KryptonConstants.BROWSER_SAUCELABS))
                    Property.RcMachineId = KryptonConstants.BROWSER_SAUCELABS;
            }

            else
            {
                Property.RcMachineId = Environment.MachineName;
                Utility.SetVariable("RCMachineId", Environment.MachineName);
                Utility.SetParameter("RCMachineId", Environment.MachineName);
            }

            if (!Environment.UserName.IsNullOrWhiteSpace())
                Property.RcUserName = Environment.UserName; //set logged in username


            if (Utility.GetParameter("FailedCountForExit").IsNullOrWhiteSpace())
            {
                Utility.SetParameter("FailedCountForExit", Property.FailedCountForExit);
            }

            if (Utility.GetParameter("StartParallelRecovery").IsNullOrWhiteSpace())
            {
                Utility.SetParameter("StartParallelRecovery", Property.StartParallelRecovery);
            }

            Property.Date_Time = Utility.GetParameter("DateTimeFormat").Replace("/", "\\/");

            //Execution start date and time set
            DateTime dtNow = DateTime.Now;

            #region source path location and clear temp folder
            var sourcePath = "KRYPTONResults-" + Guid.NewGuid();

            //HTML File location
            Property.HtmlFileLocation = string.Format("{0}{1}", Path.GetTempPath(), sourcePath);
            Property.ListOfFilesInTempFolder.Add(Property.HtmlFileLocation);
            #endregion

            Property.ExecutionStartDateTime = dtNow.ToString(Property.Date_Time);
            FilePath = new string[TestCases.Length];

            bool validation = false; //need to capture the validation steps so that it will not check for each test case

            //when running in batch mode (more than one test case or in test suite), value of this parameter is always true, irrespective of what user specified
            if (TestCases.Length > 1) Utility.SetParameter("closebrowseroncompletion", "true");

            // final test case id to be stored in final test case variable
            Property.FinalTestCase = TestCases[TestCases.Length - 1];


            #region  Setup Test Environment on Local Workstation
            for (int testCaseCnt = 0; testCaseCnt < TestCases.Length; testCaseCnt++)
            {
                // This section handles setting test environment using command line provided in Environment specific ini files
                // Retrieve command to be executed from variables
                string envSetupCommand = Utility.GetVariable("EnvironmentSetupBatch");
                envSetupCommand = Utility.ReplaceVariablesInString(envSetupCommand);

                // Do not attempt to setup environment if no script needs to be executed
                if (!(envSetupCommand.IsNullOrWhiteSpace()))
                {
                    if (File.Exists(Path.Combine(Property.EnvironmentFileLocation, envSetupCommand)))
                    {
                        Process setupProcess = new Process
                        {
                            StartInfo =
                            {
                                WindowStyle = ProcessWindowStyle.Hidden,
                                UseShellExecute = false,
                                FileName = Path.Combine(Property.EnvironmentFileLocation, envSetupCommand),
                                ErrorDialog = false,
                                Arguments = Utility.GetParameter("Browser") +
                                            " " + Utility.GetParameter("Environment") +
                                            " " + "\"" + Utility.GetParameter("TestSuite").Trim() + "\"" +
                                            " " + "\"" + Property.RcProcessId + "\""
                            }
                        };

                        //Pass on couple of extra arguments to setup batch file
                        //These includes BrowserType, TestEnvironment

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
                            KryptonException.Writeexception(e);
                        }
                    }
                    else
                    {
                        Console.WriteLine(new KryptonException("Environment file: " + Property.ApplicationPath + envSetupCommand + " is not exist."));
                        return TestSuiteResult;
                    }
                }
            #endregion
                string logDestinationExt = string.Empty;
                if (Utility.GetParameter("KeepReportHistory").Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    logDestinationExt = "-" + dtNow.ToString("ddMMyyhhmmss");
                    Property.DateTime = logDestinationExt;
                }
                Property.ResultsSourcePath = string.Format("{0}{1}\\{2}\\{3}", Path.GetTempPath(), sourcePath, TestCases[testCaseCnt].Trim(),
                                                                       (testCaseCnt + 1).ToString());


                //if user wants to keep report history, it will be available with dateandtime appended in the log destination folder
                if (!Property.RcProcessId.IsNullOrWhiteSpace())
                    logDestinationExt = "-" + Property.RcProcessId + logDestinationExt;

                if (Path.IsPathRooted(Utility.GetParameter("LogDestinationFolder")))
                    Property.ResultsDestinationPath = string.Format("{0}\\{1}{2}", Utility.GetParameter("LogDestinationFolder"),
                                                                           TestCases[testCaseCnt].Trim(), logDestinationExt);
                else
                    Property.ResultsDestinationPath = string.Format("{0}\\{1}{2}", string.Concat(Property.IniPath,
                                                                    Utility.GetParameter("LogDestinationFolder")),
                                                                           TestCases[testCaseCnt].Trim(), logDestinationExt);

                //Initialize the logwriter to record any message/exception after test step completion
                try
                {
                    Logwriter = container.Resolve<IKryptonLogger>(new ResolverOverride[] 
                    {
                        new ParameterOverride("testcasecnt",testCaseCnt)
                    });
                        //new KryptonFileLogWriter(testCaseCnt);
                }
                catch (Exception e)
                {
                    KryptonException.Writeexception(e);
                    Console.WriteLine(Exceptions.ERROR_INVALIDTESTID + TestCases[testCaseCnt] + "in Test Suite: " + TestSuite); // If test case Id will be of more than 255 characters.. 
                    return TestSuiteResult;
                }

                #region Creating the xml log file path
                try
                {
                    FilePath[testCaseCnt] = string.Format("{0}\\{1}", Property.ResultsSourcePath, Property.LogFileName);
                    Reporting.LogFile.filePathName = FilePath[testCaseCnt];
                    XmlLog = Reporting.LogFile.Instance; //xml log file initialization
                    XmlLog.AddTestAttribute("TestCase Id", TestCases[testCaseCnt].Trim());
                    if (!Property.RcMachineId.IsNullOrWhiteSpace())
                        XmlLog.AddTestAttribute("RCMachineId", Property.RcMachineId);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(new KryptonException(exception.Message));
                    return TestSuiteResult;
                }
                #endregion

                Property.ValidateSetup = Utility.GetParameter("ValidateSetup");

                if (Property.ValidateSetup.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    #region validation

                    //once it validates it will not run for each test case
                    try
                    {
                        if (validation != true)
                        {
                            Validate.GetDriverValidation(
                                Utility.GetParameter(Property.DriverString).ToLower());
                            Property.Status = Validate.validate.ValidationProcess();
                            if (Property.Status == ExecutionStatus.Fail)
                            {
                                ReportHandler.WriteExceptionLog(new KryptonException("Validation Faild"), 0, "Validation");
                                return TestSuiteResult;
                            }
                            validation = true;
                        }

                    }
                    catch (Exception exception)
                    {
                        ReportHandler.WriteExceptionLog(new KryptonException(exception.Message), 0, "Validation");
                        return TestSuiteResult;
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

                    Cursor.Position = new Point(0, 0);
                }

                _objTestManager = container.Resolve<Manager>(new ResolverOverride[] 
                { 
                    new ParameterOverride("testCaseId", TestCases[testCaseCnt].Trim()), new ParameterOverride("testCasefilename", string.Empty) 
                });

                //Call InitExecution method to allow manager pass on control 
                //to respective test manager for required settings
                try
                {
                    Manager.InitTestCaseExecution();
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
                    if (_orTestData == null)
                    {

                        if (!Path.GetExtension(Utility.GetParameter("ORFileName")).ToLower().Equals(Property.ExcelSheetExtension.ToLower()))
                            Property.ObjectRepositoryFilename = Utility.GetParameter("ORFileName").ToLower() + Property.ExcelSheetExtension;
                        else
                            Property.ObjectRepositoryFilename = Utility.GetParameter("ORFileName").ToLower();
                        Manager.SetTestFilesLocation(Property.ManagerType);  // get ManagerType from property.cs
                        OrData(); //To Fetch object Repository data and set in private global dataset

                        if (testCaseCnt == 0 && !Utility.GetParameter("RunRemoteExecution").ToLower().Equals("true"))
                        {
                            //reading RecoverPopup File.
                            xmlRecoverFromPopupFlow = _objTestManager.GetRecoverFromPopupXml();
                            xmlRecoverFromBrowserFlow = _objTestManager.GetRecoverFromBrowserXml();

                            #region Start Exe for Popups handling.
                            try
                            {
                                if (String.Compare(Utility.GetVariable("StartParallelRecovery"), "true", StringComparison.OrdinalIgnoreCase) == 0)
                                {
                                    string ext = Path.GetExtension(Property.ParallelRecoverySheetName);
                                    if (string.IsNullOrEmpty(ext))
                                        Property.ParallelRecoverySheetName = Property.ParallelRecoverySheetName + Property.ExcelSheetExtension;
                                    if (!File.Exists(Path.Combine(Property.RecoverFromPopupFilepath, Property.ParallelRecoverySheetName)))
                                    {
                                        Console.WriteLine(ConsoleMessages.MSG_PARALLEL_RECOVERY);
                                        Console.WriteLine("Please make sure  " + Property.ParallelRecoverySheetName + "  exist at  " + Property.RecoverFromPopupFilepath + "  directory.");
                                        return TestSuiteResult;
                                    }
                                    ProcessStartInfo processInfo = new ProcessStartInfo
                                    {
                                        FileName = Property.ApplicationPath + "KryptonParallelRecovery.exe",
                                        CreateNoWindow = true,
                                        WindowStyle = ProcessWindowStyle.Hidden,
                                        Arguments =
                                            "\"" + Property.RecoverFromPopupFilepath + "/" +
                                            Property.ParallelRecoverySheetName + "\"" +
                                            " " + "\"" + Property.Popup_Sheetname + "\"" +
                                            " " + Process.GetCurrentProcess().Id
                                    };

                                    Utility.SetProcessParameter("KryptonParallelRecovery.exe");
                                    Process parallelProcess = Process.Start(processInfo);
                                    Thread.Sleep(1000);
                                    if (parallelProcess.Id == 0 || parallelProcess.HasExited)
                                    {
                                        Console.WriteLine(Exceptions.ERROR_AUTOIT);
                                        Console.WriteLine(Exceptions.ERROR_RECOVERY);
                                        return TestSuiteResult;
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                //Please register AutoIt dll
                                Console.WriteLine(Exceptions.ERROR_RECOVERY + e.Message);
                                Console.WriteLine(Exceptions.ERROR_AUTOIT);
                                return TestSuiteResult;
                            }
                            #endregion
                        }
                    }
                }
                catch (Exception exception)
                {
                    ReportHandler.WriteExceptionLog(new KryptonException(exception.Message), 0, Exceptions.ERROR_OBJECTREPOSITORY);
                    return TestSuiteResult;
                }
                #endregion
                if (TestStepAction == null) //so that driver will not initialize every time
                    TestStepAction = new TestDriver.Action(xmlRecoverFromPopupFlow, xmlRecoverFromBrowserFlow, _orTestData);
                try
                {
                    DateTime dtNow1 = DateTime.Now;
                    XmlLog.AddTestAttribute("ExecutionStartTime", dtNow1.ToString(Property.Date_Time));
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    Console.WriteLine("Starting Test Case: " + TestCases[testCaseCnt].Trim());
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    _failureCount = 0;
                    _stepNos = 0; //initialize the counter
                    Property.EndExecutionFlag = false;
                    ExecuteTestCase(TestCases[testCaseCnt].Trim()); //actual test flow
                    Property.EndExecutionFlag = false;
                    //Shutting down the driver if CloseBrowserOnCompletion is set to true. 
                    if (Utility.GetParameter("closebrowseroncompletion").ToLower().Trim().Equals("true") || TestCases.Length > 1)
                        TestStepAction.Do("closeallbrowsers", "", "", "", "");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);

                    #region cleanup test flow
                    string[] cleanupCase = _cleanupTestCase.Split(',');
                    string cleanupTestCaseId = string.Empty;
                    if (cleanupCase.Length > 0 && cleanupCase[0].ToLower().Contains("cleanuptestcaseid"))
                    {
                        cleanupTestCaseId = cleanupCase[0].Split('=')[1].Replace("\"", "");
                    }

                    if (cleanupCase.Length > 1 && cleanupCase[1].ToLower().Contains("data"))
                    {
                        string[] argument = cleanupTestCaseId.Split(Property.Seprator);
                        for (int argCount = 0; argCount < argument.Length; argCount++)
                        {
                            Utility.SetVariable("argument" + (argCount + 1), argument[argCount].Trim());
                        }
                    }
                    string cleanupTestCaseIteration = null;
                    if (cleanupCase.Length > 2 && cleanupCase[2].ToLower().Contains("iteration"))
                    {
                        cleanupTestCaseIteration = cleanupCase[2].Split('=')[1].Replace("\"", ""); ;
                    }
                    if (cleanupTestCaseId.IsNullOrWhiteSpace() == false)
                        ExecuteTestCase(cleanupTestCaseId, cleanupTestCaseIteration, TestCaseId, "true");
                    #endregion
                    Console.WriteLine("Ending Test Case: " + TestCases[testCaseCnt].Trim());
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    XmlLog.AddTestAttribute("RCMachineId", Property.RcMachineId);

                    Utility.ClearRunTimeDic(); //clear run time dictionary and copy parameter dictionary in run time dictionary

                    //Calculate test execution duration
                    TimeSpan tcDurationSpan = (DateTime.Now - dtNow1);
                    TimeSpan tcDuration = new TimeSpan(tcDurationSpan.Hours, tcDurationSpan.Minutes, tcDurationSpan.Seconds);
                    XmlLog.AddTestAttribute("ExecutionDuration", tcDuration.ToString());
                    Console.WriteLine("Total test execution duration was " + tcDuration);

                    dtNow1 = DateTime.Now;
                    XmlLog.AddTestAttribute("ExecutionEndTime", dtNow1.ToString(Property.Date_Time));
                    XmlLog.AddTestAttribute("SauceJobUrl", Property.JobUrl);
                }
                catch (Exception exception)
                {
                    ReportHandler.WriteExceptionLog(new KryptonException(exception.Message), _stepNos, "Execute Test Case");
                }
                Console.WriteLine(ConsoleMessages.MSG_DASHED);
                Console.WriteLine(ConsoleMessages.MSG_EXCECUTION_COMPLETE_LOG);
                Console.WriteLine(ConsoleMessages.MSG_DASHED);

                try
                {

                    //Generation of xml log file.snapshot
                    XmlLog.SaveXmlLog();

                    //for the last step all the file will upload after generation of html file.
                    if (testCaseCnt < TestCases.Length - 1)
                        Manager.UploadTestExecutionResults();
                }
                catch (Exception ex)
                {
                    Logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
                }

                #region for Angieslist: Batch command to execution on test completion

                try
                {
                    string testFinishCommand = Utility.GetVariable("TestCleanupBatch");
                    testFinishCommand = Utility.ReplaceVariablesInString(testFinishCommand);

                    if (!(testFinishCommand.IsNullOrWhiteSpace()))
                    {
                        testFinishCommand = Path.Combine(Property.ApplicationPath, testFinishCommand);
                        if (File.Exists(testFinishCommand))
                        {
                            //Create process information and assign data
                            Process setupProcess = new Process
                            {
                                StartInfo =
                                {
                                    WindowStyle = ProcessWindowStyle.Hidden,
                                    UseShellExecute = false,
                                    FileName = testFinishCommand,
                                    WorkingDirectory = Property.ApplicationPath,
                                    ErrorDialog = false,
                                    Arguments = Property.RcProcessId +
                                                " " + "\"" + XmlLog.ExecutedXmlPath + "\""
                                }
                            };
                            // Start the process
                            setupProcess.Start();
                            //No need to wait for process to exit
                            setupProcess.WaitForExit(10000);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(Exceptions.ERROR_CLEANUP + e.Message);
                }
                #endregion
            }

            //Stoping the IE Popup Thread after execution.
            ExecuteThread = false;
            #region  Killing all the processes started by Krypton Application during execution.
            try
            {
                string[] procesesToKill = Utility.GetAllProcesses();
                foreach (string proc in procesesToKill)
                {
                    foreach (Process process in Process.GetProcessesByName(proc.Substring(0, proc.LastIndexOf('.'))))
                    {
                        try
                        {
                            process.Kill();
                            process.WaitForExit(10000);
                        }
                        catch (Exception ex)
                        {
                            Logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
            }
            #endregion

            //Shutting down the driver if CloseBrowserOnCompletion is set to true.
            if (Utility.GetParameter("Browser").Equals(KryptonConstants.BROWSER_FIREFOX, StringComparison.OrdinalIgnoreCase))
            {
            }
            if (!Utility.GetVariable("RCProcessId").IsNullOrWhiteSpace() && !Utility.GetVariable("RCProcessId").Equals("RCProcessId", StringComparison.OrdinalIgnoreCase))
                TestStepAction.Do("shutdowndriver"); //shutdown driver process running in remote

            TestDriver.Action.SaveScript();

            Console.WriteLine(ConsoleMessages.MSG_EXCECUTION_COMPLETED_HTML);

            //Execution end date and time set
            dtNow = DateTime.Now;
            Property.ExecutionEndDateTime = dtNow.ToString(Property.Date_Time);

            try
            {
                ReportHandler.CreateHtmlReportSteps(MLstInputTestCaseIDs);
            }
            catch (Exception ex)
            {
                Logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
                return TestSuiteResult;
            }

            #region  Delete temp folder.
            /*List Of files that would be deleted here are :
            1. Text file for Error message storage that would be generated per test suit.(Size in KB).
            2. WebDriver Profile Folder that would be generated per test case.(Size ~ 23MB).
            */
            try
            {
                foreach (string file in Property.ListOfFilesInTempFolder)
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
            catch (IOException ex)
            {
                Logwriter.WriteLog("IO Exception Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
            }
            catch (Exception ex)
            {
                Console.WriteLine(Utility.GetCommonMsgVariable("KRYPTONERRCODE0048"));
                Logwriter.WriteLog("Message:-" + ex.Message + " InnerException:-" + ex.InnerException + " StackTrace:-" + ex.StackTrace);
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
            using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
            {
                sw.WriteLine("Task3");
            }
            //Wait for user input at the end of the execution is handled by configuration file
            if (string.Equals(Utility.GetParameter("EndExecutionWaitRequired"), "false", StringComparison.OrdinalIgnoreCase))
            {
                Array.ForEach(Process.GetProcessesByName("cmd"), x => x.Kill());
            }
            return TestSuiteResult;
        }

        /// <summary>
        /// Set the projectpath and Set path for reading the .ini files 
        /// e.g Parameter.ini,EmailNotificaiton.ini etc
        /// </summary>
        /// <param string[]="args">Command line arguments contains the project properties.</param> 
        /// <returns></returns>
        private static void InitializeIniPath(string[] args)
        {
            try
            {
                if (File.Exists(Path.Combine(Property.ApplicationPath, "root.ini")))
                {
                    string currentProjectpath = string.Empty;
                    if (args.Length != 0)
                    {
                        for (int i = 0; i < args.Length; i++)
                        {
                            string[] keyValuePair = args[i].Trim().Split('=');
                            string key = keyValuePair[0].Trim();
                            string value = keyValuePair[1].Trim();
                            if (key.ToLower().Equals("projectpath") && !string.IsNullOrEmpty(value))
                                currentProjectpath = value;
                        }
                    }
                    if (currentProjectpath.Length == 0)
                    {
                        using (StreamReader sr = new StreamReader(Path.Combine(Property.ApplicationPath, "root.ini")))
                        {
                            string currentProjectName;
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
                            currentProjectpath = Path.Combine(Property.ApplicationPath, currentProjectpath);
                    }
                    Property.IniPath = Path.GetFullPath(currentProjectpath);
                }
                else
                {
                    Property.IniPath = Property.ApplicationPath;
                }

            }
            catch (Exception e)
            {
                throw new Exception(Exceptions.ERROR_INVALIDCOMMANDLINE);
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
            DataSet xmlTestCaseDataSet = null;
            int testCaseIdExists = -1;
            string projectSpecificTestCaseFile = string.Empty;
            Utility.SetVariable("reusabledefaultfile", testCaseId.Substring(0, testCaseId.IndexOf('.')));
            #region Retrieve test case workflow from test manager and set to local dataset
            DataSet xmlTestCaseFlow;
            try
            {
                if (projectSpecific == "true")
                {
                    #region project specific reusable test cases fetching
                    string globalActionFileName = "";
                    if (!Path.GetExtension(Utility.GetVariable("reusabledefaultfile").ToLower().Trim()).Equals(Property.ExcelSheetExtension.ToLower()))
                        globalActionFileName = Utility.GetVariable("reusabledefaultfile") + Property.ExcelSheetExtension;
                    else
                        globalActionFileName = Utility.GetVariable("reusabledefaultfile").Trim();

                    string TempglobalActionFileName = globalActionFileName.Substring(0, globalActionFileName.LastIndexOf('.'));
                    if (Property.TestCaseIdParameter == 0)
                        projectSpecificTestCaseFile = Common.Utility.GetParameter("TestCaseId");
                    else
                        projectSpecificTestCaseFile = Property.TestCaseIdSeperator + TempglobalActionFileName;

                    _objTestManager = container.Resolve<Manager>(new ResolverOverride[] 
                    {
                        new ParameterOverride("testCaseId", projectSpecificTestCaseFile), new ParameterOverride("testCasefilename", globalActionFileName) 
                    });
                        //new Krypton.Manager(projectSpecificTestCaseFile, globalActionFileName);
                    xmlTestCaseFlow = _objTestManager.GetTestCaseXml(Property.TestQuery,true,false);

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
                        string[] sep = { Property.TestCaseIdSeperator};
                        string[] testFlowSheetName = testCaseId.Split(sep, StringSplitOptions.None);
                        xmlTestCaseFlow = _objTestManager.GetTestCaseXml(testFlowSheetName[0], false, true);
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
                    _objTestManager = container.Resolve<Manager>(new ResolverOverride[]
                    {
                        new ParameterOverride("testCaseId", testCaseId), new ParameterOverride("testCasefilename", string.Empty) 
                    });
                        //new Manager(testCaseId);
                    xmlTestCaseFlow = _objTestManager.GetTestCaseXml(Property.TestQuery, false, false);

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
                using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
                {
                    sw.WriteLine("In catch throw 0049exception 3");
                }
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
                Utility.SetVariable("Iteration", testDataCounter.ToString());
                if (isTestDataUsed)
                {
                    xmlTestCaseDataSet = null;
                }

                //to check if the loop reach the starting of the current test case records.
                bool currentTestCaseRecords = false;
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
                        XmlLog.AddTestAttribute("TestCase Name", xmlTestCaseName);
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
                    xmlTestCaseStepAction = Utility.ReplaceVariablesInString(xmlTestCaseStepAction);
                    xmlTestCaseData = Utility.ReplaceVariablesInString(xmlTestCaseData);
                    xmlTestCaseParent = Utility.ReplaceVariablesInString(xmlTestCaseParent);
                    xmlTestCaseObj = Utility.ReplaceVariablesInString(xmlTestCaseObj);
                    xmlTestCaseOption = Utility.ReplaceVariablesInString(xmlTestCaseOption);
                    xmlTestCaseComment = Utility.ReplaceVariablesInString(xmlTestCaseComment);
                    xmlTestCaseIteration = Utility.ReplaceVariablesInString(xmlTestCaseIteration);
                    #endregion

                    #region Read individual column values from Test Case Flow and pass on to test driver
                    if (xmlTestCaseStepAction.IsNullOrWhiteSpace() == false && ValidStepOption(xmlTestCaseOption) && Property.EndExecutionFlag != true)
                    {
                        string cleanUpString = Utility.GetCleanupTestCase(xmlTestCaseOption, "cleanuptestcaseid"); //capture Clean up test case id with options

                        if (!cleanUpString.Equals(string.Empty))
                            _cleanupTestCase = cleanUpString;

                        if (xmlTestCaseStepAction.Equals("runExcelTestCase", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            //iteration method call for internal test case
                            ExecuteTestCase(xmlTestCaseData, xmlTestCaseIteration, testCaseId, null);
                        }
                        else
                        {
                            string testCaseData;
                            if (isTestDataUsed)
                            {

                                if (xmlTestCaseDataSet == null) // will call test data method only once
                                {
                                    #region call to test case data method to fetch corresponding data
                                    try
                                    {
                                        //fetching test data sheet name from the option field in the first line of the nested test case and set for the whole test case 
                                        string[] testDataSheet1 = Utility.GetTestMode(xmlTestCaseOption,
                                                                                          "TestDataSheet").Split(',');
                                        var testDataSheet = testDataSheet1.Length > 1 ? testDataSheet1[1] : "test_data";
                                        Property.TestDataSheet = testDataSheet;

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
                                        _stepNos++;
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
                                    //to fetch by [td] variable
                                    testCaseData = "{$[TD]" + xmlTestCaseObj + "}";
                                }
                            }

                            else
                                testCaseData = xmlTestCaseData;

                            #region Handling test data situations for verifyDatabase
                            if (xmlTestCaseStepAction.Equals("verifyDatabase", StringComparison.OrdinalIgnoreCase))
                            {
                                //for data from DBTestData sheet
                                try
                                {
                                    DbTestData(testCaseId, xmlTestCaseData, xmlTestCaseObj);
                                    testCaseData = xmlTestCaseData;
                                }
                                catch (Exception exception)
                                {
                                    _stepNos++;
                                    throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0050"), exception.Message);
                                }
                            }
                            #endregion


                            testCaseData = testCaseData.Trim();

                            #region Starting special script based on data given in options column.
                            string[] scriptToRunContents = Utility.GetTestMode(xmlTestCaseOption, "script").Split(',');
                            if (scriptToRunContents.Length > 1)
                            {
                                var scriptToRun = scriptToRunContents[1].ToLower();
                                try
                                {

                                    if (!(scriptToRun.Contains(".exe")))  // Updated to Check if script provided already contains an extension
                                    {
                                        scriptToRun = scriptToRun + ".exe";
                                    }
                                    Console.WriteLine("Executing special script: " + scriptToRun);

                                    Utility.SetProcessParameter(scriptToRun); //Add process name to process array before starting it.
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
                                catch (Exception ex)
                                {
                                    KryptonException.Writeexception(ex);
                                }
                            }

                            #endregion
                            //Replace variables in string here for all columns and pass on
                            xmlTestCaseStepAction = Utility.ReplaceVariablesInString(xmlTestCaseStepAction);
                            xmlTestCaseParent = Utility.ReplaceVariablesInString(xmlTestCaseParent);
                            xmlTestCaseObj = Utility.ReplaceVariablesInString(xmlTestCaseObj);
                            xmlTestCaseOption = Utility.ReplaceVariablesInString(xmlTestCaseOption);
                            xmlTestCaseComment = Utility.ReplaceVariablesInString(xmlTestCaseComment);
                            xmlTestCaseIteration = Utility.ReplaceVariablesInString(xmlTestCaseIteration);
                            testCaseData = Utility.ReplaceVariablesInString(testCaseData);

                            if (xmlTestCaseOption.IndexOf("{append}", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                testCaseData = Utility.ReplaceVariablesInString("{$" + xmlTestCaseObj + "}") + testCaseData;
                            }


                            //Replace variable with special charecters)

                            if (ValidStepOption(testCaseData) == false) continue;

                            #region call to TestDriver Action.Do method
                            try
                            {
                                //Initialize Step Log.
                                Property.InitializeStepLog();
                                if (xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0 && !(xmlTestCaseStepAction.ToLower().Equals("setvariable") && Property.DebugMode.Equals("false")))
                                {
                                    _stepNos++;
                                    Property.StepNumber = _stepNos.ToString();
                                }
                                Console.WriteLine("Parent:" + xmlTestCaseParent);
                                Console.WriteLine("Object:" + xmlTestCaseObj);
                                Console.WriteLine("Step Action:" + xmlTestCaseStepAction);
                                if (testCaseData.IndexOf("{$[TD]", StringComparison.OrdinalIgnoreCase) < 0)
                                    Console.WriteLine("Data:" + testCaseData);
                                else Console.WriteLine("Data:");

                                //Set TestDriver object dictionary
                                TestDriver.Action.ObjDataRow = GetTestOrData(xmlTestCaseParent, xmlTestCaseObj);

                                // In case of DragAndDrop, there are two objects to be handled simultaneously.
                                // Set second object dictionary 
                                if (xmlTestCaseStepAction.Equals("DragAndDrop", StringComparison.OrdinalIgnoreCase) == true)
                                    TestDriver.Action.ObjSecondDataRow = GetTestOrData(xmlTestCaseParent, xmlTestCaseData);

                                if (TestDriver.Action.ObjDataRow.ContainsKey(KryptonConstants.WHAT))
                                    Console.WriteLine("OR Locator:" + TestDriver.Action.ObjDataRow[KryptonConstants.WHAT]);
                                else
                                    Console.WriteLine("OR Locator:");

                                Property.ExecutionDate = DateTime.Now.ToString(Utility.GetParameter("DateFormat"));
                                Property.ExecutionTime = DateTime.Now.ToString(Utility.GetParameter("TimeFormat"));

                                TestStepAction.Do(xmlTestCaseStepAction, xmlTestCaseParent, xmlTestCaseObj,
                                                   testCaseData,
                                                   xmlTestCaseOption);

                                if (xmlTestCaseStepAction.Equals("closeallbrowsers",StringComparison.OrdinalIgnoreCase))
                                {
                                    Utility.SetVariable("FirefoxProfilePath", Utility.GetParameter("FirefoxProfilePath"));
                                }

                                //if action method is not implemented is returned, need to check for project specific functions
                                if (Property.Remarks.IndexOf(Utility.GetCommonMsgVariable("KRYPTONERRCODE0025"), StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    #region checkprojectspecificfunction
                                    if (xmlTestCaseStepAction.IndexOf(Property.TestCaseIdSeperator,
                                                                      StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        if (xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0)
                                        {
                                            _stepNos--;
                                        }
                                        #region KRYPTON0252
                                        string[] argument = testCaseData.Split(Property.Seprator);
                                        if (argument.Length > 0)
                                        {
                                            for (int argCount = 0; argCount < argument.Length; argCount++)
                                            {
                                                Utility.SetVariable("argument" + (argCount + 1), argument[argCount].Trim());
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
                                    bool flag = xmlTestCaseOption.Equals("{NoReporting}") && Property.Status == ExecutionStatus.Fail || xmlTestCaseOption.Equals("{NoReporting}").Equals(false);
                                    if ((xmlTestCaseOption.IndexOf("{optional}", StringComparison.OrdinalIgnoreCase) < 0))
                                    {
                                        if (flag)
                                        {
                                            if (Property.Status == ExecutionStatus.Fail &&
                                                xmlTestCaseStepAction.IndexOf("verify",
                                                                              StringComparison.OrdinalIgnoreCase) < 0)
                                            {
                                                TestSuiteResult = 1; //process fail
                                                //Increment Failure Count only in case of non optional and non ignored step cases.

                                                _failureCount++;

                                                if ((_failureCount ==
                                                    int.Parse(Utility.GetParameter("FailedCountForExit")) &&
                                                    int.Parse(Utility.GetParameter("FailedCountForExit")) > 0)
                                                    || Property.EndExecutionFlag == true)
                                                {
                                                    if (Property.EndExecutionFlag == true)
                                                        Console.WriteLine(ConsoleMessages.MSG_EXCEUTION_ENDS);
                                                    else
                                                        Console.WriteLine(ConsoleMessages.MSG_FAILURE_EXCEED);

                                                    Property.StepComments = xmlTestCaseComment +
                                                        ": End execution due to nos. of failure(s) exceeds failed counter defined in parameter file -" + _failureCount + ".";
                                                    XmlLog.WriteExecutionLog();

                                                    Property.EndExecutionFlag = true;
                                                    return;
                                                }
                                                Property.StepComments = xmlTestCaseComment;
                                                XmlLog.WriteExecutionLog();
                                                Console.WriteLine("Status:" + Property.Status);
                                                Console.WriteLine("Remarks:" + Property.Remarks);
                                            }
                                            else
                                            {
                                                Property.StepComments = xmlTestCaseComment;
                                                if (!(xmlTestCaseStepAction.ToLower().Equals("setvariable") && Property.DebugMode.Equals("false")))
                                                    XmlLog.WriteExecutionLog();
                                                Console.WriteLine("Status:" + Property.Status);
                                                Console.WriteLine("Remarks:" + Property.Remarks);
                                            }
                                        }
                                    }
                                    //set browser attribute to xml log file and that will be call once
                                    if (xmlTestCaseBrowser == null && !Utility.GetVariable(Property.BrowserVersion).Equals("version", StringComparison.OrdinalIgnoreCase))
                                    {
                                        xmlTestCaseBrowser = Utility.GetVariable(Property.BrowserString) + Utility.GetVariable(Property.BrowserVersion);
                                        XmlLog.AddTestAttribute("Browser", Utility.GetVariable(Property.BrowserString) + Utility.GetVariable(Property.BrowserVersion));
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
                                XmlLog.SaveXmlLog(); //Generation of xml log file 
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
            DataSet xmlTestCaseDataSet;
            {
                string globalActionFileName;
                if (!Path.GetExtension(Utility.GetParameter("reusabledefaultfile").ToLower().Trim()).Equals(Property.ExcelSheetExtension.ToLower()))
                    globalActionFileName = Utility.GetParameter("reusabledefaultfile") + Property.ExcelSheetExtension;
                else
                    globalActionFileName = Utility.GetParameter("reusabledefaultfile").Trim();

                Manager objTestManager = null;

                objTestManager = projectspecific ? container.Resolve<Manager>(new ResolverOverride[]
                                   {
                                       new ParameterOverride("testCaseId", childTestCaseId), new ParameterOverride("testCasefilename", globalActionFileName)
                                   })
                    : container.Resolve<Manager>(new ResolverOverride[]
                                   {
                                       new ParameterOverride("testCaseId", childTestCaseId), new ParameterOverride("testCasefilename", string.Empty)
                                   });

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
                            Utility.SetVariable(columnName, dataValue); //set value
                        }
                    }
                }

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
        private static void DbTestData(string testCaseId, string rowNo, string testObject)
        {
            Manager objTestManager = container.Resolve<Manager>(new ResolverOverride[]
                                   {
                                       new ParameterOverride("testCaseId", testCaseId), new ParameterOverride("testCasefilename", string.Empty)
                                   });
            _xmlTestDataSet = objTestManager.GetDbTestDataXml();
            //Set test data in the runtime dictionary for parameter calls in the test case by {$[TD]##$}
            for (int dsDataCnt = 0; dsDataCnt < _xmlTestDataSet.Tables[0].Rows.Count; dsDataCnt++)
            {
                if (_xmlTestDataSet.Tables[0].Rows[dsDataCnt]["SNo"].ToString().Equals(rowNo))
                {
                    for (int dsColumnCnt = 0;
                        dsColumnCnt < _xmlTestDataSet.Tables[0].Columns.Count;
                        dsColumnCnt++)
                    {
                        string columnName = "[TD]" +
                                            _xmlTestDataSet.Tables[0].Columns[dsColumnCnt].ColumnName;
                        //get Column name
                        string dataValue = _xmlTestDataSet.Tables[0].Rows[dsDataCnt][dsColumnCnt].ToString();
                        //get data
                        Utility.SetVariable(columnName, dataValue);
                    }
                }
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

                foreach (DataRow drData in _orTestData.Tables[0].Rows)
                {

                    if (drData[KryptonConstants.PARENT].ToString().Trim().Equals(parent, StringComparison.OrdinalIgnoreCase)
                        && drData[KryptonConstants.TEST_OBJECT].ToString().Trim().Equals(testObj, StringComparison.OrdinalIgnoreCase)
                        && drData[KryptonConstants.TEST_OBJECT].ToString().Trim().IsNullOrWhiteSpace() == false)
                    {
                        if (orDataRow.ContainsKey(KryptonConstants.LOGICAL_NAME) == false)
                        {
                            orDataRow.Add(KryptonConstants.PARENT, Utility.ReplaceVariablesInString(drData[KryptonConstants.PARENT].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.LOGICAL_NAME, Utility.ReplaceVariablesInString(drData[KryptonConstants.LOGICAL_NAME].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.OBJ_TYPE, Utility.ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.HOW, Utility.ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.WHAT, Utility.ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.MAPPING, Utility.ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.TEST_OBJECT, Utility.ReplaceVariablesInString(drData[KryptonConstants.TEST_OBJECT].ToString().Trim()));//  to retrieve object in the HTML also.(PBI 118 in TFs)
                        }
                        else
                        {
                            orDataRow[KryptonConstants.PARENT] = Utility.ReplaceVariablesInString(drData[KryptonConstants.PARENT].ToString().Trim());
                            orDataRow[KryptonConstants.LOGICAL_NAME] = Utility.ReplaceVariablesInString(drData[KryptonConstants.LOGICAL_NAME].ToString().Trim());
                            orDataRow[KryptonConstants.OBJ_TYPE] = Utility.ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString().Trim());
                            orDataRow[KryptonConstants.HOW] = Utility.ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString().Trim());
                            orDataRow[KryptonConstants.WHAT] = Utility.ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString().Trim());
                            orDataRow[KryptonConstants.MAPPING] = Utility.ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString().Trim());
                            orDataRow[KryptonConstants.TEST_OBJECT] = Utility.ReplaceVariablesInString(drData[KryptonConstants.TEST_OBJECT].ToString().Trim());// to retrieve object in the HTML also.(PBI 118 in TFs)
                        }
                    }

                }
            }
            catch (KryptonException)
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
            var dsOr = Manager.GetObjectDefinitionXml();
            _orTestData = dsOr;
            Property.OrBachupData = dsOr;
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
            bool validStep = !((stepOption.IndexOf("{ignore}", StringComparison.OrdinalIgnoreCase) >= 0 && stepOption.IndexOf("{ignorespace}", StringComparison.OrdinalIgnoreCase) < 0 && stepOption.IndexOf("{ignorecase}", StringComparison.OrdinalIgnoreCase) < 0) ||
                stepOption.IndexOf("{skip}", StringComparison.OrdinalIgnoreCase) >= 0);

            string[] testmode = Utility.GetTestMode(stepOption, Property.TestMode).Split(',');

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