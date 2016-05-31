/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: TestManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: File to Interact With TFS.
*****************************************************************************/
using System;
using System.Linq;
using System.Diagnostics;
using System.IO;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Common;

namespace Krypton
{
    public class TestManager
    {
        //Declare all class level variables here
        static readonly string ProjectUrl = Utility.GetParameter("TFSUrl");
        static readonly string ProjectName = Utility.GetParameter("TFSProjectName");
        static readonly string TestPlanName = Utility.GetParameter("TestPlanName");
        static readonly string TestSuiteName = Property.CurrentTestCase.Split('.').First();
        static readonly int TestCaseId = Convert.ToInt32(Property.CurrentTestCase.Split('.').Last());
        static readonly string TestSuiteId = Utility.GetParameter("TestSuiteId");
        static readonly string MhtReportGenerationExe = Utility.GetParameter("MHTGenerationExeName");

        /// <summary>
        /// This method is used to initiate execution in test manager, create test run and results and set required parameters
        /// </summary>
        /// <returns></returns>
        public bool InitExecution()
        {
            //Create a connection to tfs project
            ITestManagementTeamProject tfsProject = null;
            tfsProject = GetProject(ProjectUrl, ProjectName);

            if (tfsProject == null)
            {
                throw new Exception("Unabled to connect to test project: " + ProjectName);
            }
            //Retrieve test plan details
            ITestPlanCollection testPlans = tfsProject.TestPlans.Query("select * from TestPlan where PlanName ='" +
                                                                       TestPlanName + "'");
            if (testPlans.Count == 0)
            {
                throw new Exception("Unabled to locate test plan: " + TestPlanName + " in Test Manager.");
                  
            }

            ITestPlan tfsTestPlan = testPlans.First();

            //Retrieve test suite details
            ITestSuiteCollection testSuites = null;

            //Optionally, test suite id of test manager can be passed as an command line arguments
            //This helps when same test case has been added to multiple test suites
            if (TestSuiteId.ToLower().Equals(string.Empty) ||
                TestSuiteId.ToLower().Equals(string.Empty) ||
                TestSuiteId.ToLower().Equals("testsuiteid", StringComparison.OrdinalIgnoreCase))
            {
                testSuites = tfsProject.TestSuites.Query("Select * from TestSuite where Title='" +
                                                         TestSuiteName + "' and PlanID='" + tfsTestPlan.Id + "'");
            }
            else
            {
                testSuites = tfsProject.TestSuites.Query("Select * from TestSuite where Id='" +
                                                         TestSuiteId + "' and PlanID='" + tfsTestPlan.Id + "'");
            }


            IStaticTestSuite tfsTestSuite = testSuites.Cast<IStaticTestSuite>().FirstOrDefault(testSuite => testSuite.Title.ToLower().Equals(TestSuiteName.ToLower()) || testSuite.Id.ToString().Equals(TestSuiteId));

            if (tfsTestSuite == null)
            {
                throw new Exception("Unabled to locate test suite: " + TestSuiteName + " in Test Manager Test Plan: " + TestPlanName);
                   
            }

            //Get handle to a specific test case in the test suite
            ITestCase tfsTestCase = tfsTestSuite.AllTestCases.FirstOrDefault(testcase => testcase.Id.Equals(TestCaseId));

            if (tfsTestCase == null)
            {
                throw new Exception("Unabled to locate test case id: " + TestCaseId + " in Test Manager");
            }

            //Create a test run
            ITestPoint tfsTestPoint = CreateTestPoints(tfsTestPlan, tfsTestSuite, tfsTestCase);
            ITestRun tfsTestRun = CreateTestRun(tfsProject, tfsTestPlan, tfsTestPoint);
            tfsTestRun.Refresh();

            //Suprisingly, most recently created test results should be available in last, but test manager returns it at first position
            //Find test results that were create by the test run
            ITestCaseResultCollection tfsTestCaseResults = tfsProject.TestResults.ByTestId(tfsTestCase.Id);
            ITestCaseResult tfsTestResult = tfsTestCaseResults.Last();  //Default assignment
            foreach (ITestCaseResult testResult in tfsTestCaseResults)
            {
                if (testResult.DateCreated.CompareTo(tfsTestRun.DateCreated) == 1)
                {
                    tfsTestResult = testResult;
                    break;
                }
            }

            //Set test run and result id to property variable for usage while uploading results
            Property.RcTestRunId = tfsTestRun.Id;
            Property.RcTestResultId = tfsTestResult.TestResultId;

            //Set status of test case execution

            //Set other details on test execution
            tfsTestResult.ComputerName = Property.RcMachineId;
            tfsTestResult.DateStarted = DateTime.Now;
            tfsTestResult.State = TestResultState.InProgress;
            tfsTestResult.Save();
            return true;
        }


        public string DownloadAttachment(string attachmentLocation, string strFileName)
        {
            return attachmentLocation + "/" + strFileName;
        }

        /// <summary>
        /// This method will upload test results to test lab in visual studio test manager
        /// </summary>
        /// <param name="xmlFileList">
        /// List of all test results xml files, saparate by semicolumn. 
        /// All other files are expected to be present in this location
        /// </param>
        /// <param name="testCaseStatus"></param>
        /// Status of executed test case, can be pass, fail or warning
        /// <returns></returns>
        ///  
        public static bool UploadTestResults(string xmlFileList, string testCaseStatus) 
        {
            Array xmlFiles = xmlFileList.Split(';');
            ITestManagementTeamProject tfsProject = null;
            tfsProject = GetProject(ProjectUrl, ProjectName);

            if (tfsProject == null)
            {
                throw new Exception("Unabled to connect to test project: " + ProjectName);
            }
            //Above section commented as this is not required when InitExecution method is implemented
            ITestCaseResult tfsTestResult = null;
            tfsTestResult = tfsProject.TestResults.Find(Property.RcTestRunId, Property.RcTestResultId);

            if (tfsTestResult == null)
            {
                throw new Exception("Unabled to locate test results in Test Manager");
            }

            //Set status of test case execution
            switch (testCaseStatus.ToLower())
            {
                case "fail":
                case "failed":
                    tfsTestResult.Outcome = TestOutcome.Failed;
                    break;
                case "pass":
                case "passed":
                    tfsTestResult.Outcome = TestOutcome.Passed;
                    break;
                case "warning":
                    tfsTestResult.Outcome = TestOutcome.Warning;
                    break;
            }
            tfsTestResult.Save();
            //Attach files
            foreach (string xmlFileName in xmlFiles)
            {
                DirectoryInfo resultLocation = new DirectoryInfo(xmlFileName).Parent;

                //Generate MHT report
                try
                {
                    string iconLocation = new DirectoryInfo(Property.ApplicationPath).FullName;
                    iconLocation = iconLocation.TrimEnd('\\');
                    Process generateMht = new Process();
                    generateMht.StartInfo.FileName = Property.ApplicationPath + MhtReportGenerationExe;
                    if (resultLocation != null)
                        generateMht.StartInfo.Arguments = "\"" + resultLocation.FullName + "\"" + " " + "\"" + iconLocation + "\"";
                    generateMht.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    generateMht.Start();
                    generateMht.WaitForExit();
                    generateMht.Close();
                }
                catch
                {
                    //No through
                }
                //Traverse through each file
                if (resultLocation != null)
                    foreach (FileInfo resultFile in resultLocation.GetFiles())
                    {
                        if (!(resultFile.Extension.ToLower().Contains("jpeg") ||
                              resultFile.Extension.ToLower().Contains("html") ||
                              resultFile.Extension.ToLower().Contains("jpg")))
                        {
                            ITestAttachment resultAttachment = tfsTestResult.CreateAttachment(resultFile.FullName);
                            tfsTestResult.Attachments.Add(resultAttachment);
                            tfsTestResult.Save();
                        }
                    }
            }

            //Set other details on test execution
            try
            {
                tfsTestResult.RunBy = tfsTestResult.GetTestRun().Owner;
                tfsTestResult.DateCompleted = DateTime.Now;
                tfsTestResult.Duration = DateTime.Now - tfsTestResult.DateStarted;
                tfsTestResult.ErrorMessage = Property.ExecutionFailReason;
                tfsTestResult.State = TestResultState.Completed;
                tfsTestResult.Save();
            }
            catch
            {
                // ignored
            }
            tfsTestResult.Save();
            return true;
        }

        public static ITestManagementTeamProject GetProject(string serverUrl, string project)
        {
            try
            {
                TfsTeamProjectCollection tfs = new TfsTeamProjectCollection(TfsTeamProjectCollection.GetFullyQualifiedUriForName(serverUrl));
                ITestManagementService tms = tfs.GetService<ITestManagementService>();

                return tms.GetTeamProject(project);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        //create test point
        public static ITestPoint CreateTestPoints(ITestPlan testPlan, ITestSuiteBase suite, ITestCase testcase)
        {
            try
            {
                ITestPointCollection tpc = testPlan.QueryTestPoints("SELECT * FROM TestPoint WHERE SuiteId = " + suite.Id + " and TestCaseID =" + testcase.Id);
                ITestPoint tp = testPlan.FindTestPoint(tpc[0].Id);
                return tp;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //create test run
        private static ITestRun CreateTestRun(ITestManagementTeamProject project, ITestPlan plan, ITestPoint tp)
        {
            try
            {
                ITestRun run = plan.CreateTestRun(true);
                run.AddTestPoint(tp, null);
                run.Save();
                return run;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static void DeleteTestResults(ITestCaseResultCollection allTestResults, int olderThanDays)
        {
            try
            {
                foreach (ITestCaseResult testResult in allTestResults)
                {
                    if ((DateTime.Now - testResult.DateCreated).TotalDays > olderThanDays)
                    {
                        try
                        {
                            testResult.GetTestRun().Delete();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(ConsoleMessages.EXCEPTION_WHILE_DELETING_TEST_RUNS + e.Message);
            }
        }

    }
}