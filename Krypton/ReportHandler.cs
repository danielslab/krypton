using System;
using System.Collections.Generic;
using Common;
using Reporting;

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
        public static void WriteExceptionLog(KryptonException exception, int stepNo, string stepDescription)
        {
            Console.WriteLine(exception.Message);
            try
            {
                Property.InitializeStepLog();
            }
            catch (Exception ex)
            {
                KryptonException.ReportException(ex.Message+"Initializesteplog()");
            }
            Property.StepNumber = stepNo.ToString();
            Property.StepDescription = stepDescription;
            Property.Status = ExecutionStatus.Fail;
            Property.Remarks = exception.Message;
            Property.ExecutionDate = DateTime.Now.ToString(Utility.GetParameter("DateFormat"));
            Property.ExecutionTime = DateTime.Now.ToString(Utility.GetParameter("TimeFormat"));
            try
            {
                XmlLog.WriteExecutionLog();
                XmlLog.SaveXmlLog();
            }
            catch (Exception e) 
            {
                KryptonException.ReportException(e.Message + "--->" + e.StackTrace + "--->" + e.Source);    
            }
            
            if (Utility.GetParameter("closebrowseroncompletion").ToLower().Trim().Equals("true"))
            {
                if (TestStepAction != null)
                    TestStepAction.Do("closeallbrowsers");
            }

            try
            {
                if (stepDescription.IndexOf("Execute Test Case", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    //Execution end date and time set
                    DateTime dtNow = DateTime.Now;
                    Property.ExecutionEndDateTime = dtNow.ToString(Property.Date_Time);

                    CreateHtmlReportSteps();

                    if (Utility.GetParameter("closebrowseroncompletion").ToLower().Trim().Equals("true"))
                    {
                        if (TestStepAction != null && !Utility.GetParameter("RunRemoteExecution").Equals("true", StringComparison.OrdinalIgnoreCase))
                            TestStepAction.Do("shutdowndriver"); //shutdown driver
                    }
                }
            }
            catch(Exception ex){
                TestEngine.Logwriter.WriteLog("Data" + ex.Data + "Stacktrace" + ex.StackTrace + "Message" + ex.Message);
            }
            //Wait for user input at the end of the execution is handled by configuration file
            if (!string.Equals(Utility.GetParameter("EndExecutionWaitRequired"), "false", StringComparison.OrdinalIgnoreCase)
                && stepDescription.IndexOf("Execute Test Case", StringComparison.OrdinalIgnoreCase) < 0)
            {
                Console.WriteLine(ConsoleMessages.MSG_DASHED);
                Console.WriteLine(ConsoleMessages.MSG_TERMINATE_WITH_EXCEPTION);
                while (true)
                {
                    ConsoleKeyInfo inf = Console.ReadKey(true); // Key output not shown
                    if (inf.Key == ConsoleKey.Enter) break;
                        
                    Console.WriteLine("Please press [Enter]" + inf.Key);
                }
            }
            TestSuiteResult = 1;
        }

        /// <summary>
        /// common steps for creating html report and upload
        /// </summary>
        internal static void CreateHtmlReportSteps(List<string> inputTestIds = null)
        {
            //Creating HTML Report
            try
            {
                LogFile.AllXmlFilesLocation = string.Join(";", FilePath);

                // Using new format of report.
                bool isSummaryRequired = Utility.GetParameter("SummaryReportRequired").ToLower().Equals("true");
                bool htmlReportRequired = !Utility.GetParameter("HtmlReportRequired").ToLower().Equals("false");
                LogFile.CreateHtmlReport(string.Empty, false, true, Property.IsSauceLabExecution);
                HtmlReport.CreateHtmlReport(string.Empty, false, true, Property.IsSauceLabExecution, isSummaryRequired, false, false, htmlReportRequired);
                GetReportSummary();
                Property.FinalXmlPath = FilePath[FilePath.Length - 1];
            }
            catch (Exception exception)
            {
                TestSuiteResult = 1;
                KryptonException.ReportException(exception.Data +"\n Message " +exception.Message +"\n StackTrace"+ exception.StackTrace+"\n Line Number"+
                    exception.StackTrace.Substring(exception.StackTrace.LastIndexOf(' ')));
            }

            Console.WriteLine(ConsoleMessages.MSG_DASHED);
            Console.WriteLine(ConsoleMessages.MSG_UPLOADING_LOG);

            Manager.UploadTestExecutionResults();
            try
            {
                if (Utility.GetParameter("EmailNotification").Equals("true", StringComparison.OrdinalIgnoreCase))
                    Utility.EmailNotification("end", false); //Email Notification after completion of the process
            }
            catch (Exception)
            {
                Console.WriteLine(Utility.GetCommonMsgVariable("KRYPTONERRCODE0053"));
                throw;
            }
        }


        //This method is creating report summary
        private static void GetReportSummary()
        {
            return;
            int startLocation = 0;
            string strRepotSummary = string.Empty;
            string text_to_replace = string.Empty;
            strRepotSummary = Property.ReportSummaryBody;
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
            Property.ReportSummaryBody = strRepotSummary;
        }

    }
}
