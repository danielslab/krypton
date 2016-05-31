/****************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Reporting.LogFile.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Creating Xml and Html Reports
*****************************************************************************/
using T= System;
using System;
using System.Collections.Generic;
using System.Collections;
using System.Data;


namespace Common
{

    public class Property
    {
        
        public static string ProductName = "Krypton";
        public static string ApplicationPath = string.Empty; 
        public static string IniPath=string.Empty; 
        public static string ParameterFileName = "Parameters.ini";
        public static char KeyValueDistinctionKeyword = ':';
        public static Dictionary<string, string> Parameterdic = new Dictionary<string, string>();
        public static Dictionary<string, string> Runtimedic = new Dictionary<string, string>();
        public static Dictionary<string, string> CommonMsgdic = new Dictionary<string, string>();
        public static ArrayList ProcessLists = new ArrayList(); //Contains all the processes that have been started during execution.

        public static string JavaKeyString = @"SOFTWARE\JavaSoft";
        public static string JavaRunTimeEnvironmentKeyString = @"SOFTWARE\JavaSoft\Java Runtime Environment";
        public static string BrowserKeyString = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
        public static string BrowserKeyString64Bit = @"Software\Wow6432Node\Mozilla";
        public static string ChromeSpecificKeyString = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths";
        public static string JavaRunTimeEnvironmentString = "Java Runtime Environment";
        public static string BrowserString = "browser"; //Keyword for specifying browser in parameter.ini file.
        public static string BrowserVersion = "version";    //Keyword to store browser's version. Set from Browser.cs
        public static string IeProcess = "iexplore";
        public static string FfProcess = "firefox";
        public static string ChromeProcess = "chrome";
        public static string Selenium = "selenium";
        public static string Qtp = "qtp";
        public static string DriverString = "driver";
        public static string Environment = "Environment";
        public static string KryptonVersion = "Krypton 2.0.0";
        public static int RecoveryCount = 30;
        public static bool IsRecoveryRunning = false;
        public static string DestinationFolderDownload = @"\results";
        public static List<string> ListOfFilesInTempFolder = new List<string>();
        public static DataSet OrBachupData = new DataSet();
        public static string TempiniFileName = "temp";
        
        
        public static String ReportSummaryBody = string.Empty; 
        public static int TotalTestCaseExecutionTime=0;
        public static string TestCaseFilepath = "test_cases"; //temporary to check
        public static string ReusableLocation = "Reusables";
        public static string TestSuiteFileLocation = @".\TestSuites";
        public static string ParallelRecoverySheetName = "popuprecovery";
        public static string EnvironmentFileLocation=@".\TestEnvironment";
        public static int Waitforalert = 30;    
   
        public static string TestQuery = @"select * from [test_flow$] ";
        
        public static string Popup_Sheetname = "IE_Popups";
        public static string TestDataFilepath = "test_cases"; //temporary to check
        public static string TestDataSheet = string.Empty;
        public static string RemoteMachineIP = string.Empty;
         

        
        public static string DBTestDataFilepath = "test_cases"; //temporary to check
        public static string RecoverFromPopupFilepath = "recovery_scenarios"; //temporary
        public static string RecoverFromBrowserFilePath = "recovery_scenarios";//temporary
        public static string ObjectRepositoryFilepath = "object_repository"; //temporary
        public static string ObjectRepositoryFilename = "ObjectRepository";
        
        public static string RecoverFromPopupFilename = "RecoveryFromPopups";
        public static string RecoverFromBrowserFilename = "RecoveryFromBrowser";
        public static string ExcelSheetExtension = ".csv";

        public static string ResultsSourcePath = string.Empty;

        public static string ErrorLog = @"C:\Krypton\ErrorLog\Logs.txt";
        public static string ResultsDestinationPath = @"C:\Krypton\ResultsDestination";
        public static char Seprator = '|';
        public static string TestMode = "TestMode";
        //currently using any arbitary path.
        public static string Time = "hh:mm:ss tt";
        public static string Date = "dd/MM/yyyy";
        public static string Date_Time = "yyyy_MM_dd_HH_mm_ss";
        public static string StepActionResult = string.Empty;
        public static string DateTime = string.Empty;
        public static string StepNumber = string.Empty;
        public static string StepDescription = string.Empty;
        public static string Status = string.Empty;
        public static string ExecutionDate = string.Empty;
        public static string ExecutionTime = string.Empty;
        public static string Remarks = string.Empty;
        public static string Attachments = string.Empty;
        //This property will contain url of browser when snapshot was taken
        public static string AttachmentsUrl = string.Empty;

        public static string ObjectHighlight = string.Empty;
        public static string StepComments = string.Empty;

        public static string ExecutionStartDateTime = string.Empty;
        public static string ExecutionEndDateTime = string.Empty;
        public static string FinalExecutionStatus = string.Empty;
        public static string JobExecutionStatus = ExecutionStatus.Pass; 
        public static string JobUrl = string.Empty; 
        public static bool IsSauceLabExecution = false;
        public static string HtmlFileLocation = string.Empty;
        public static string CompanyLogo = string.Empty;
        public static string GetRequest = "get";
        public static string PostRequest = "post";

        //Name of API Response file where server response would be saved. Extenstion is determined by reponseformat specified in test data
        public static string ApiXmlFile = "ApiResponse";
        public static string ApiXmlHeaderFile = "ResponseHeader.header";
        public static string DownloadedFileName = "WebFile";

        //LogFileName by default set to executionlog.xml
        public static string LogFileName = "executionlog.xml";

        public static bool NoWait = false;
		 
        #region Public Fields Declartion code for Sauce Labs
        /// <summary>
        ///  These field will be used for validating/Accessing Sauce Labs parameters
        /// </summary>
        public static string SauceLabs = "saucelabs";
        public static string SauceLabsParameterFile = "saucelabs.ini";
        #endregion
        
        public static void InitializeStepLog()
        {
            StepDescription = string.Empty;
            Status = string.Empty;
            ExecutionDate = string.Empty;
            ExecutionTime = string.Empty;
            Remarks = string.Empty;
            Attachments = string.Empty;
            ObjectHighlight = string.Empty;
            StepComments = string.Empty;
            HtmlSourceAttachment = string.Empty;
            AttachmentsUrl = string.Empty;
        }

        //Database connection string initialization to fetch data from database
        public static string DbConnectionString = string.Empty;

        public static string ErrorCaptureAs = string.Empty;
        //using dummy server for database methods.
        public static string SqlConnectionStringRead = "server=Kartik-D10\\SuperExpress;Trusted_Connection=yes;connection timeout=120;Database=SampleDB;Network Library=DBMSSOCN";
        public static string SqlConnectionStringWrite = "server=Kartik-D10\\SuperExpress;Trusted_Connection=yes;connection timeout=120;Database=SampleDB;Network Library=DBMSSOCN";
        //using arbitary path for sql query files.
        public static string SqlQueryFilePath = @"./SQLQueries";

        public static string TestCaseIdSeperator = ".";

        public static int TestCaseIdParameter = 0;

        //email notification template path
        public static string EmailNotificationFile = "EmailNotification.ini";
        public static string EmailStartTemplate = @"templates\StartTemplate.txt";
        public static string EmailEndTemplate = @"templates\EndTemplate.txt";


        public static string SnapshotOption = string.Empty;
        public static string DebugMode = string.Empty;
        public static string RemoteUrl = string.Empty;
        public static string IsRemoteExecution = string.Empty;

        public static string ReportSettingsFile = "ReportSettings.ini";
        public static string ReportZipFileName = "Report.zip";

        public static string HtmlSourceAttachment = string.Empty;

        public static string ValidateSetup = "true";

        public static string GlobalTimeOut = "30";

        public static string GenerateRandomNumberParamName = "randomnumber"; 
        public static string GenerateRandomStringParamName = "UniqueString";

        //Default Url specified in ini file :
        public static string ApplicationUrl = string.Empty;

        // Stores windows handle when launching a new browser
        public static IEnumerable<string> ArrKnownBrowserHwnd;
        // This will store handle of original window
        public static string HwndFirstWindow = string.Empty;
        // This will store handle of most recently known windows
        public static string HwndMostRecentWindow = string.Empty; 

        // Compression ratio for images taken, to reduce size of images. 
        //25L means 25%, 50L will mean 50% reduction in quality.
        public static long ImageCompressionRatio = 75L;

        public static int MaxTimeoutForPageLoad = 0;
        public static int MinTimeoutForPageLoad = 0;

        //Added to handle end of execution from test case
        public static bool EndExecutionFlag = false; 

        public static string GlobalCommentsFile = "CommonMessages.kry";
        public static string ScriptLanguage = "php";

        public static string ManagerType = "FileSystem";   
        public static int TotalStepExecuted = 0;
        public static int TotalStepPass = 0;
        public static int TotalStepFail = 0;
        public static int TotalStepWarning = 0;
        public static int TotalCaseExecuted = 0;
        public static int TotalCasePass = 0;
        public static int TotalCaseFail = 0;
        public static int TotalCaseWarning = 0;


        public static string FailedCountForExit = "0";

        public static string StartParallelRecovery = "False";
        public static string[] ListOfUniqueCharacters = {
                                                            "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l",
                                                            "m", "n", "o", "p", "q",
                                                            "r", "s", "t", "u", "v", "w", "x", "y", "z"
                                                        };

        public static string RcProcessId = string.Empty;
        public static string RcMachineId = "localhost";
        public static string FinalXmlPath = string.Empty;

        //final test case id to be stored in final test case variable
        public static string FinalTestCase = string.Empty;

        //Currently executing test case if
        public static string CurrentTestCase = string.Empty;

        public static string ExecutionFailReason = string.Empty; //final test case failure reason to be stored

        public static string RcUserName = string.Empty; //logged in username

        //Test Run Id Property, set by test manager for current run
        public static int RcTestRunId = -1;

        //Test Results Id Property, set by test manager for current run
        public static int RcTestResultId = -1;
   
        public static string CurrentTime = T.DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".txt";
        public static string LogFolder = "ErrorLog";

        public static string ParallelRecoveryFilePath = string.Empty;
    }



    public static class ExecutionStatus
    {
        public static readonly string Pass = "Pass";
        public static readonly string Fail = "Fail";
        public static readonly string Warning = "Warning";
    }

}
