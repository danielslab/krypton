using System.Collections.Generic;

namespace Krypton
{
    class Variables
    {
        protected static Reporting.LogFile XmlLog = null; //private  null declation of log file
        protected static TestDriver.Action TestStepAction = null; //Delacre TestDriver Action class
        protected static int TestSuiteResult = 0; //private  variable to set complete test suite as pass or fail
        protected static string[] FilePath = null; //private  variable for all the test cases file paths
        protected static string TestCaseId = string.Empty;
        protected static string TestSuite = string.Empty;
        protected static string[] TestCases = null;
        protected static List<string> MLstInputTestCaseIDs = new List<string>();
    }
}
