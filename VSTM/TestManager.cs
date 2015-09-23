/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: TestManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: File to Interact With TFS.
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.XPath;

namespace Krypton
{
    public class TestManager
    {
        //Declare all class level variables here
        static string ProjectUrl = Common.Utility.GetParameter("TFSUrl");
        static string projectName = Common.Utility.GetParameter("TFSProjectName");
        static string testPlanName = Common.Utility.GetParameter("TestPlanName");
        static string testSuiteName = Common.Property.CurrentTestCase.Split('.').First();
        static int testCaseId = Convert.ToInt32(Common.Property.CurrentTestCase.Split('.').Last());
        static string testSuiteId = Common.Utility.GetParameter("TestSuiteId");
        static string MHTReportGenerationExe = Common.Utility.GetParameter("MHTGenerationExeName");

        /// <summary>
        /// This method is used to initiate execution in test manager, create test run and results and set required parameters
        /// </summary>
        /// <returns></returns>
        public bool InitExecution()
        {
            try
            {
                //Create a connection to tfs project
                ITestManagementTeamProject tfsProject = null;
                tfsProject = GetProject(ProjectUrl, projectName);

                if (tfsProject == null)
                {
                    throw new Exception("Unabled to connect to test project: " + projectName);
                }
                //Retrieve test plan details
                ITestPlanCollection testPlans = tfsProject.TestPlans.Query("select * from TestPlan where PlanName ='" +
                                                testPlanName + "'");
                if (testPlans.Count() == 0)
                {
                    throw new Exception("Unabled to locate test plan: " + testPlanName + " in Test Manager.");
                  
                }

                ITestPlan tfsTestPlan = testPlans.First();

                //Retrieve test suite details
                ITestSuiteCollection testSuites = null;
                IStaticTestSuite tfsTestSuite = null;

                //Optionally, test suite id of test manager can be passed as an command line arguments
                //This helps when same test case has been added to multiple test suites
                if (testSuiteId.ToLower().Equals(string.Empty) ||
                    testSuiteId.ToLower().Equals(string.Empty) ||
                    testSuiteId.ToLower().Equals("testsuiteid", StringComparison.OrdinalIgnoreCase))
                {
                    testSuites = tfsProject.TestSuites.Query("Select * from TestSuite where Title='" +
                                                      testSuiteName + "' and PlanID='" + tfsTestPlan.Id + "'");
                }
                else
                {
                    testSuites = tfsProject.TestSuites.Query("Select * from TestSuite where Id='" +
                                                      testSuiteId + "' and PlanID='" + tfsTestPlan.Id + "'");
                }


                foreach (IStaticTestSuite testSuite in testSuites)
                {
                    if (testSuite.Title.ToLower().Equals(testSuiteName.ToLower()) ||
                        testSuite.Id.ToString().Equals(testSuiteId))
                    {
                        tfsTestSuite = testSuite;
                        break;
                    }
                }

                if (tfsTestSuite == null)
                {
                    throw new Exception("Unabled to locate test suite: " + testSuiteName + " in Test Manager Test Plan: " + testPlanName);
                   
                }

                //Get handle to a specific test case in the test suite
                ITestCase tfsTestCase = null;
                foreach (ITestCase testcase in tfsTestSuite.AllTestCases)
                {
                    if (testcase.Id.Equals(testCaseId))
                    {
                        tfsTestCase = testcase;
                        break;
                    }
                }

                if (tfsTestCase == null)
                {
                    throw new Exception("Unabled to locate test case id: " + testCaseId + " in Test Manager");
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
                Common.Property.RCTestRunId = tfsTestRun.Id;
                Common.Property.RCTestResultId = tfsTestResult.TestResultId;

                //Set status of test case execution

                //Set other details on test execution
                tfsTestResult.ComputerName = Common.Property.RCMachineId;
                tfsTestResult.DateStarted = DateTime.Now;
                tfsTestResult.State = TestResultState.InProgress;
                tfsTestResult.Save();

            }
            catch (Exception exception)
            {
                throw exception;
            }
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
            try
            {
                 Array xmlFiles = xmlFileList.Split(';');
                ITestManagementTeamProject tfsProject = null;
                tfsProject = GetProject(ProjectUrl, projectName);

                if (tfsProject == null)
                {
                    throw new Exception("Unabled to connect to test project: " + projectName);
                }
                //Above section commented as this is not required when InitExecution method is implemented
                ITestCaseResult tfsTestResult = null;
                tfsTestResult = tfsProject.TestResults.Find(Common.Property.RCTestRunId, Common.Property.RCTestResultId);

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
                        string iconLocation = new DirectoryInfo(Common.Property.ApplicationPath).FullName;
                        iconLocation = iconLocation.TrimEnd('\\');
                        Process generateMHT = new Process();
                        generateMHT.StartInfo.FileName = Common.Property.ApplicationPath + MHTReportGenerationExe;
                        generateMHT.StartInfo.Arguments = "\"" + resultLocation.FullName + "\"" + " " + "\"" + iconLocation + "\"";
                        generateMHT.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        generateMHT.Start();
                        generateMHT.WaitForExit();
                        generateMHT.Close();
                    }
                    catch
                    {
                        //No through
                    }
                    //Traverse through each file
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
                    tfsTestResult.ErrorMessage = Common.Property.ExecutionFailReason;
                    tfsTestResult.State = TestResultState.Completed;
                    tfsTestResult.Save();
                }
                catch
                {
                }
                tfsTestResult.Save();
            }
            catch (Exception exception)
            {
                throw exception;
            }
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
                Console.WriteLine("Exception while trying to delete older test runs. Error: " + e.Message);
            }
        }

    }
}