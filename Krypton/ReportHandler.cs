using System;
using System.Collections.Generic;
using System.Linq;
using Common;
using Reporting;
using System.Text;

namespace Krypton
{
    class ReportHandler : Variables
    {
        /// <summary>
        /// This is the method to write common exception to xml log and display in console
        /// </summary>
        /// <param name="exception">Exception</param>
        /// <param name="stepNo">Step Number</param>
        /// <param name="stepDescription">Description</param>
        /// <returns></returns>
        public static void WriteExceptionLog(Common.KryptonException exception, int stepNo, string stepDescription)
        {
            Console.WriteLine(exception.Message);
            Common.Property.InitializeStepLog();
            Common.Property.StepNumber = stepNo.ToString();
            Common.Property.StepDescription = stepDescription;
            Common.Property.Status = Common.ExecutionStatus.Fail;
            Common.Property.Remarks = exception.Message;
            Common.Property.ExecutionDate = DateTime.Now.ToString(Common.Utility.GetParameter("DateFormat"));
            Common.Property.ExecutionTime = DateTime.Now.ToString(Common.Utility.GetParameter("TimeFormat"));

            xmlLog.WriteExecutionLog();//Generation of xml log file
            xmlLog.SaveXmlLog();
            if (Utility.GetParameter("closebrowseroncompletion").ToLower().Trim().Equals("true"))
            {
                if (testStepAction != null)
                    testStepAction.Do("closeallbrowsers");

            }

            try
            {
                if (stepDescription.IndexOf("Execute Test Case", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    //Execution end date and time set
                    DateTime dtNow = DateTime.Now;
                    Common.Property.ExecutionEndDateTime = dtNow.ToString(Common.Property.DATE_TIME);

                    CreateHtmlReportSteps();

                    if (Utility.GetParameter("closebrowseroncompletion").ToLower().Trim().Equals("true"))
                    {
                        if (testStepAction != null && !Utility.GetParameter("RunRemoteExecution").Equals("true", StringComparison.OrdinalIgnoreCase))
                            testStepAction.Do("shutdowndriver"); //shutdown driver

                    }
                }
            }
            catch
            {

            }
            //Wait for user input at the end of the execution is handled by configuration file
            if (!string.Equals(Common.Utility.GetParameter("EndExecutionWaitRequired"), "false", StringComparison.OrdinalIgnoreCase)
                && stepDescription.IndexOf("Execute Test Case", StringComparison.OrdinalIgnoreCase) < 0)
            {
                Console.WriteLine(ConsoleMessages.MSG_DASHED);
                Console.WriteLine(ConsoleMessages.MSG_TERMINATE_WITH_EXCEPTION);
                while (true)
                {
                    ConsoleKeyInfo inf = Console.ReadKey(true); // Key output not shown
                    if (inf.Key == ConsoleKey.Enter) break;
                    else
                    {
                        Console.WriteLine("Please press [Enter]" + inf.Key);
                    }
                }
            }
            testSuiteResult = 1;
            return;
        }

        /// <summary>
        /// common steps for creating html report and upload
        /// </summary>
        internal static void CreateHtmlReportSteps(List<string> InputTestIds = null)
        {
            //Creating HTML Report
            try
            {
                LogFile.allXmlFilesLocation = string.Join(";", filePath);

                // Using new format of report.
                bool IsSummaryRequired = Common.Utility.GetParameter("SummaryReportRequired").ToLower().Equals("true");
                bool HtmlReportRequired = Utility.GetParameter("HtmlReportRequired").ToLower().Equals("false") ? false : true;
                Reporting.LogFile.CreateHtmlReport(string.Empty, false, true, Property.isSauceLabExecution);
                Reporting.HTMLReport.CreateHtmlReport(string.Empty, false, true, Property.isSauceLabExecution, IsSummaryRequired, false, false, HtmlReportRequired);
                GetReportSummary();
                Property.finalXmlPath = filePath[filePath.Length - 1]; //assign the last xml result file path

            }
            catch (Exception exception)
            {
                testSuiteResult = 1;
                throw exception;
            }

            Console.WriteLine(ConsoleMessages.MSG_DASHED);
            Console.WriteLine(ConsoleMessages.MSG_UPLOADING_LOG);

            Manager.UploadTestExecutionResults(InputTestIds); //upload log file to filesystem/QC/XStudio
            try
            {
                if (Common.Utility.GetParameter("EmailNotification").Equals("true", StringComparison.OrdinalIgnoreCase))
                    Common.Utility.EmailNotification("end", false); //Email Notification after completion of the process
            }
            catch (Exception exception)
            {
                Console.WriteLine(Utility.GetCommonMsgVariable("KRYPTONERRCODE0053"));
                throw exception;
            }
        }


        //This method is creating report summary
        private static void GetReportSummary()
        {
            return;
            int startLocation = 0;
            string strRepotSummary = string.Empty;
            string text_to_replace = string.Empty;
            strRepotSummary = Common.Property.ReportSummaryBody;
            //Remove detailed steps from the report
            int i = 0;
            int endlocation = 0;
            do
            {
                i = i + 1;
                startLocation = strRepotSummary.IndexOf("<div id='tblID", 0);
                if (startLocation <= 0)
                    break;
                endlocation = strRepotSummary.IndexOf("</div>", startLocation) + 6;
                text_to_replace = strRepotSummary.Substring(startLocation, endlocation - startLocation);
                strRepotSummary = strRepotSummary.Replace(text_to_replace, "");
                strRepotSummary = strRepotSummary.Replace("style='table-layout:fixed'", "");
            }
            while (i < 500);

            int intCounter;
            int expandLocation;
            // Add serial number to the summary report instead of expand link
            intCounter = 1;
            do
            {
                expandLocation = strRepotSummary.IndexOf(">[+]<", 0);
                if (expandLocation <= 0)
                    break;
                strRepotSummary = strRepotSummary.Replace(">[+]<", ">" + intCounter.ToString() + "<");
                intCounter = intCounter + 1;
            }
            while (intCounter < 500);

            // 'Remove all links
            strRepotSummary = strRepotSummary.Replace("<a>", string.Empty);
            strRepotSummary = strRepotSummary.Replace("<A>", string.Empty);
            strRepotSummary = strRepotSummary.Replace("<\a>", string.Empty);
            strRepotSummary = strRepotSummary.Replace(@"<\A>", string.Empty);
            string totalTestCaseString;
            //'Extract total  test case count
            totalTestCaseString = "Test Case(s) Executed:";
            startLocation = strRepotSummary.IndexOf(totalTestCaseString, 0);
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            endlocation = strRepotSummary.IndexOf("<", startLocation);

            string totalExecuted;
            int inttotalExecuted;
            totalExecuted = strRepotSummary.Substring(startLocation, endlocation - startLocation);
            inttotalExecuted = int.Parse(totalExecuted);
            //Extract passed test case count
            string totalPassedString;
            totalPassedString = "Total Passed:";
            startLocation = strRepotSummary.IndexOf(totalPassedString, 0);
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            endlocation = strRepotSummary.IndexOf("<", startLocation);

            string totalPassed;
            int inttotalPassed;
            totalPassed = strRepotSummary.Substring(startLocation, endlocation - startLocation);
            inttotalPassed = int.Parse(totalPassed);

            //Extract failed test case count
            string totalFailedString;
            totalFailedString = "Total Failed:";
            startLocation = strRepotSummary.IndexOf(totalFailedString, 0);
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            endlocation = strRepotSummary.IndexOf("<", startLocation);

            string totalFailed;
            int inttotalFailed;
            totalFailed = strRepotSummary.Substring(startLocation, endlocation - startLocation);
            inttotalFailed = int.Parse(totalFailed);

            //Extract warning test case count
            string totalWarningString;
            totalWarningString = "Total Warning:";
            startLocation = strRepotSummary.IndexOf(totalWarningString, 0);
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            startLocation = strRepotSummary.IndexOf(">", startLocation) + 1;
            endlocation = strRepotSummary.IndexOf("<", startLocation);
            string totalWarning;
            int inttotalWarning;
            totalWarning = strRepotSummary.Substring(startLocation, endlocation - startLocation);
            inttotalWarning = int.Parse(totalWarning);
            inttotalPassed = inttotalPassed + inttotalWarning;
            Common.Property.ReportSummaryBody = strRepotSummary;
        }

    }
}
