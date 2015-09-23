using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Krypton
{
    class Variables
    {
        protected static Reporting.LogFile xmlLog = null; //private  null declation of log file
        protected static TestDriver.Action testStepAction = null; //Delacre TestDriver Action class
        protected static int testSuiteResult = 0; //private  variable to set complete test suite as pass or fail
        protected static string[] filePath = null; //private  variable for all the test cases file paths
        protected static string testCaseId = string.Empty;
        protected static string testSuite = string.Empty;
        protected static string[] testCases = null;
        protected static List<string> m_lstInputTestCaseIDs = new List<string>();
    }
}
