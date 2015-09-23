/****************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Reporting.LogFile.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Creating Xml and Html Reports
** **************************************************************************/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Data;
using Common;
using System.IO.Compression;
using Ionic.Zip;


namespace Reporting
{

    public class LogFile
    {
        // datatable by which Xml will be created
         private static DataTable testCaseTable;
        // string which contains Xml file path
        public string executedXmlPath = string.Empty;
        public static string allXmlFilesLocation = string.Empty;//
        public static TimeSpan timeTaken = new TimeSpan(0, 0, 0);
        public static Boolean  firstEntryInLogFile=true;
        public string executionTime = null;
        public LogFile(string filePathName)
        {
            testCaseTable = new DataTable("TestStep");
            // Copy the file name -- it will be used in function executeStep
            executedXmlPath = filePathName;
            // Storing all Xml Path file names at one place (will be used while generating HTML report)
            allXmlFilesLocation = allXmlFilesLocation + "," + filePathName;

            // Rows Structure
            testCaseTable.Columns.Add("StepNumber", typeof(string));
            testCaseTable.Columns.Add("StepDescription", typeof(string));
            testCaseTable.Columns.Add("Status", typeof(string));
            testCaseTable.Columns.Add("ExecutionDate", typeof(string));
            testCaseTable.Columns.Add("ExecutionTime", typeof(string));
            testCaseTable.Columns.Add("Remarks", typeof(string));
            testCaseTable.Columns.Add("Attachments", typeof(string));
            testCaseTable.Columns.Add("HtmlAttachments", typeof(string));
            testCaseTable.Columns.Add("ObjectHighlight", typeof(string));
            testCaseTable.Columns.Add("StepComments", typeof(string));
            testCaseTable.Columns.Add("AttachmentUrl", typeof(string));


            //creating the folder structure here
            string directory = Path.GetDirectoryName(filePathName);
            CreateDirectory(new DirectoryInfo(directory));


        }

        /// <summary>
        ///Creating folders and sub-folders. It is called iteratively 
        /// till a already present directory is not found.
        /// </summary>
        /// <param name="directory">This will contain Folder name and if not present 
        /// then called again and in end sub child folders are formed one by one</param>
        /// <returns></returns>
        private static void CreateDirectory(DirectoryInfo directory)
        {
            if (!directory.Parent.Exists)
                CreateDirectory(directory.Parent);

            //for deleting already exists files 
            if (directory.Exists)
            {
                string[] filePaths = Directory.GetFiles(directory.ToString());
                foreach (string filePath in filePaths)
                    File.Delete(filePath);
            }
            else
            {
                directory.Create();
            }
        }


        // Adding each step details
        public void WriteExecutionLog()
        {

            // Taking Parameters from Property File
            string stepNumber = Property.StepNumber.ToString();
            string stepDescription = Property.StepDescription.ToString();
            string status = Property.Status.ToString();
            string executionDate = Property.ExecutionDate.ToString();
          //Updated Code to display accurate Time in seconds and milliseconds of each stepAction 
            if (firstEntryInLogFile)
            {
                executionTime = String.Format("{0:00.00}", (DateTime.Now.TimeOfDay - Convert.ToDateTime(Property.ExecutionTime).TimeOfDay).TotalSeconds); //only first time in log file
                firstEntryInLogFile = false;
            }
            else
            {
                executionTime = String.Format("{0:00.00}", (DateTime.Now.TimeOfDay - timeTaken).TotalSeconds);

            }
             timeTaken = DateTime.Now.TimeOfDay;            
         
            string remarks = Property.Remarks.ToString();
            string attachments = Property.Attachments.ToString();
            string attachmentUrl = Property.AttachmentsUrl.ToString();
            string htmlAttachments = Property.HtmlSourceAttachment.ToString();
            string objectHighlight = Property.ObjectHighlight.ToString();
            string stepComments = Property.StepComments.ToString();
            if (status.Equals(ExecutionStatus.Fail))
                Property.JobExecutionStatus = status;

            // Creating a new row for table and pupulating row's cells
            DataRow dr = testCaseTable.NewRow();
            dr["StepNumber"] = stepNumber;
            dr["StepDescription"] = stepDescription;
            dr["Status"] = status;
            dr["ExecutionTime"] = executionTime;           
            dr["Remarks"] = remarks;
            if (attachments != string.Empty) // when some image is present
            {
                DirectoryInfo drInfo = new DirectoryInfo(executedXmlPath);
                if (drInfo.Parent != null)
                {
                    dr["Attachments"] = drInfo.Parent.Parent.Name + @"\" + drInfo.Parent.Name + @"\" + attachments;
                    dr["AttachmentUrl"] = attachmentUrl;
                }
            }
            else
            {
                dr["Attachments"] = string.Empty;
                dr["AttachmentUrl"] = string.Empty;
            }
            if (htmlAttachments != string.Empty)
            {
                DirectoryInfo drInfo = new DirectoryInfo(executedXmlPath);
                if (drInfo.Parent != null)
                {
                    dr["HtmlAttachments"] = drInfo.Parent.Parent.Name + @"\" + drInfo.Parent.Name + @"\" + htmlAttachments;
                }
            }
            else
            {
                dr["HtmlAttachments"] = "";
            }
            dr["ObjectHighlight"] = objectHighlight;
            dr["StepComments"] = stepComments;
            int failCounter = 0;
            if (status.Equals(ExecutionStatus.Fail))
            {
                Common.Property.JobExecutionStatus = ExecutionStatus.Fail;
                failCounter++;
                if (failCounter.Equals(1))
                    Property.ExecutionFailReason = remarks;
            }
            // adding newly created row to table
            testCaseTable.Rows.Add(dr);

        }

        public void SaveXmlLog()
        {
            //creating a new Xml file (replacing old Xml with new one)
            testCaseTable.WriteXml(executedXmlPath, XmlWriteMode.WriteSchema);
        }

        /// <summary>
        ///For adding new attributes to datatable (used in creating Xml Report)
        /// </summary>
        /// <param></param>
        /// <returns></returns> 
        public void AddTestAttribute(string attribute, string value)
        {
            testCaseTable.ExtendedProperties[attribute] = value;

        }

        /// <summary>
        ///Creating Html Report from various XMLs
        /// </summary>
        /// <param></param>
        /// <returns></returns>    
        public void CreateHtmlReport1()
        {
            try
            {
                DateTime executionStartTime = DateTime.ParseExact(Property.ExecutionStartDateTime, Common.Utility.GetParameter("DateTimeFormat"),
                                                                  null);
                DateTime executionEndTime = DateTime.ParseExact(Property.ExecutionEndDateTime, Common.Utility.GetParameter("DateTimeFormat"),
                                                                null);
                string generalLogoLocation = "logo" + Path.GetExtension(Property.CompanyLogo); 
                string passColor = Utility.GetParameter("ReportPassResultFontColor");
                string failColor = Utility.GetParameter("ReportFailResultFontColor");
                string warningColor = Utility.GetParameter("ReportWarningResultFontColor");
                string browserDetails = Common.Utility.GetParameter(Common.Property.BrowserString); 
                string browserVersion = Common.Utility.GetVariable(Property.BrowserVersion);
                string[] xmlFilesLocation = allXmlFilesLocation.Split(';');
                DataSet dsStore = new DataSet();
                string tcName = string.Empty;
                string tcID = string.Empty;
                string tcBrowser = string.Empty;

                string companyLogo = Property.CompanyLogo;
                string htmlFileLocation = Common.Property.HtmlFileLocation + "/HtmlReport.html";
                string finalStatus = ExecutionStatus.Pass;
                Property.FinalExecutionStatus = ExecutionStatus.Pass;
                string finalStatusColor = Utility.GetParameter("ReportPassResultFontColor");
                string populatedTr = string.Empty;
                string finalHtml = string.Empty;
                string temp = string.Empty;
                string tblID = string.Empty;
                string expander = string.Empty;
                string temp2 = string.Empty;
                string htmlStart =
                    @"<html>
                                    <head><style>td {overflow:hidden;}</style>
                                           <script>
                                                 function doMenu(item,obj) {
                                                                            tbl=item;
                                                                            if (tbl.style.display=='none') 
                                                                              {
                                                                                 tbl.style.display='block'; 
                                                                                 obj.innerHTML= '[-]';
                                                                               }
                                                                             else {
                                                                                  tbl.style.display='none';
                                                                                  obj.innerHTML= '[+]';
                                                                                  }
                                                                             }

                                            </script>
                                        </head>
                                            <body>
                                                 <div>                                                    
                                                    <p>
                                                     <!--<font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Browser:</font>
                                                      <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @";  color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" + browserDetails + 
                    @"</font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Browser Version:</font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @";  color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" + browserVersion +
                    @"</font>&nbsp;&nbsp;&nbsp;-->
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Execution Started at: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    executionStartTime.ToString(Common.Utility.GetParameter("DateTimeFormat")) +
                    @"</font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Execution Finished at: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    executionEndTime.ToString(Common.Utility.GetParameter("DateTimeFormat")) +
                    @"</font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Environment: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" + browserVersion +
                    Common.Utility.GetParameter(Common.Property.ENVIRONMENT) + 
                    @"</font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Executed by: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    "1.0.0.0" +
                    @"</font></p>
                                                    <p><font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Run Duration: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") +
                    @";'>$totalRunDuration</font></p>
                    <p>
                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Test Case(s) Executed:</font>
                                                      <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @";  color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCaseExecuted</font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Passed:</font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @";  color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCasePass 
                    </font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Failed: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCaseFail 
                    </font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Warning: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCaseWarning
                    </font>
                    </p>
                                                    <p>
                                                    </p></div><p>
                                              <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Execution Report</font>
                                              <table width='100%' style='border:1px solid " +
                    Utility.GetParameter("ReportTableBorderColor") +
                    @"'>
                                                <td>
                                                    <table  id='header'  width='100%' style='table-layout:fixed'>
                                                         <tr style='font-weight: " +
                    Utility.GetParameter("ReportTableHeaderFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportTableHeaderFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportTableHeaderFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportTableHeaderFontColor") + @"; background-color: " +
                    Utility.GetParameter("ReportTableHeaderBackgroundColor") +
                    @"'>
                                                         <td width='2%'></td>
                                                         <td width='35%'>Steps</td>                                                         
                                                         <td width='10%'>Status</td>
                                                         <td width='10%'>Browser</td>
                                                         <td width='15%'>Execution Time</td>
                                                         <td width='28%'>Remarks</td>
                                                        
                                                    </table>";


                // added remarks in each test case header
                //                     Added date/time of execution in each test case header
                string newTestCaseBodyStart = @"<table width='100%' style='table-layout:fixed'>
                                            <tr style='font-family: " + Utility.GetParameter("ReportFont") + @"; background-color: " + Utility.GetParameter("ReportTestCaseBackgroundColor") + @"'>
                                             <td width='2%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'><a id='$xn' href='javascript:doMenu($TblID,$xn)'>[+]</a></td>
                                             <td width='35%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>[$TCId] $TCName</td>                                             
                                             <td width='10%' style='font-weight: " + Utility.GetParameter("ReportTestCaseStatusWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; '><span id='span_1' style='color: $FinalStatusColor;'>$finalStatus</span></td>
                                            <td width='10%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_Browser</td>
                                            <td width='15%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_ExecutionTime_Date</td>
                                             <td width='28%' style='font-weight: " + Utility.GetParameter("ReportRemarksFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportRemarksFontStyle") + @"; color: " + Utility.GetParameter("ReportRemarksFontColor") + @"'>$Header_TC_Remarks</td>
                                            
                                            </tr>
                                           </table>
                                           <div id='$TblID' style='display:none'>";

                // Introduction of one more column (only Description)
                string trStructure = @"<table width='100%' style='word-wrap: break-word; table-layout: fixed;'> 
                                        <tr style='font-family: " + Utility.GetParameter("ReportFont") + @"; background-color: $trBackground;'>
                                            <td width='2%' ></td>
                                            <td width='35%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$stepNo<!--<font style='font-style:italic'>-->$TC_Comments<!--</font>-->$TC_Des_only</td>                                             
                                            <td width='10%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; '>$TC_Status $ViewHtml</td> 
                                            <td width='10%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>&nbsp;</td>
                                            <td width='15%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$TC_ExecutionTime_Date</td>
                                            <td width='28%' style='font-weight: " + Utility.GetParameter("ReportRemarksFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportRemarksFontStyle") + @"; color: " + Utility.GetParameter("ReportRemarksFontColor") + @"'>$TC_Remarks</td>                                                                                           
                                        </tr>
                                     </table>";
                string newTestCaseBodyEnd = @"</div>";
                string htmlEnd = @"</td>
                                 </tr>
                                  </table>
                                            </body>
                                </html>";

                //

                // Html Start
                finalHtml = htmlStart;

                int count = 0;


                foreach (string xmlLocation in xmlFilesLocation)
                {
                    if (File.Exists(xmlLocation))
                    {
                        Property.TotalCaseExecuted++;

                        try
                        {
                            #region Individual TestCase

                            dsStore = new DataSet();
                            // copying all data from XML to datatable
                            dsStore.ReadXml(xmlLocation);
                            DataTable dtTemp = new DataTable();

                            dtTemp = dsStore.Tables[0];

                            // getting TestCase Name and ID from extended properties
                            if (dtTemp.ExtendedProperties["TestCase Name"] != null)
                                tcName = dtTemp.ExtendedProperties["TestCase Name"].ToString(); 
                            if (dtTemp.ExtendedProperties["TestCase Id"] != null)
                                tcID = dtTemp.ExtendedProperties["TestCase Id"].ToString(); 
                            if (dtTemp.ExtendedProperties["Browser"] != null)
                                tcBrowser = dtTemp.ExtendedProperties["Browser"].ToString(); 

                            tblID = "tblID_" + count;
                            expander = "x_" + count;

                            temp = newTestCaseBodyStart;
                            temp = temp.Replace("$TCName", tcName);
                            temp = temp.Replace("$TCId", tcID);
                            temp = temp.Replace("$TblID", tblID);
                            temp = temp.Replace("$xn", expander);
                            // Html Start + Expand/Collapse sign (for each TestCase)
                            finalHtml = finalHtml + temp;

                            // Calculating time duration for each TestCase
                            DateTime startTime, endTime;
                            startTime = Convert.ToDateTime(dtTemp.Rows[0][4].ToString());
                            endTime = Convert.ToDateTime(dtTemp.Rows[dtTemp.Rows.Count - 1][4].ToString());

                            #region Each TestCase Step

                            int rowNo = 0;

                            //  flag to keep track of first warning/failed step.
                            bool firstFail = true;
                            bool firstWarning = true;
                            //  headerRemark is empty if test case is passed. Else it will be updated on a failed or warning step.
                            string headerRemarkFail = string.Empty;
                            string headerRemarkWarning = string.Empty;

                            //  added start date/time in test case header
                            finalHtml = finalHtml.Replace("$Header_TC_ExecutionTime_Date", dtTemp.Rows[0][3].ToString() + " : " + dtTemp.Rows[0][4].ToString());

                            //update browser info
                            finalHtml = finalHtml.Replace("$Header_TC_Browser", tcBrowser);


                            foreach (DataRow tcRow in dtTemp.Rows) // Loop over the rows.
                            {
                                Common.Property.TotalStepExecuted++;

                                // taking basic TR structure
                                populatedTr = trStructure;

                                // populating row with data (present in Xml row)
                                string stpnmb = tcRow.ItemArray[0].ToString();
                                string stpdesp = tcRow.ItemArray[1].ToString();
                                string stats = tcRow.ItemArray[2].ToString();
                                string exedt = tcRow.ItemArray[3].ToString();
                                string exetm = tcRow.ItemArray[4].ToString();
                                string rmrk = tcRow.ItemArray[5].ToString();
                                string attch = tcRow.ItemArray[6].ToString();
                                string htmlattch = tcRow.ItemArray[7].ToString();
                                string objhigh = tcRow.ItemArray[8].ToString();
                                string com = tcRow.ItemArray[9].ToString();
                                int outParam = 0;
                                Math.DivRem(rowNo, 2, out outParam);
                                if (outParam > 0)
                                {
                                    populatedTr = populatedTr.Replace("$trBackground",
                                                                      Utility.GetParameter(
                                                                          "ReportTableAlternateRowColor"));
                                }
                                else
                                {
                                    populatedTr = populatedTr.Replace("$trBackground",
                                                                      Utility.GetParameter("ReportTableRowColor"));
                                }



                                temp = "[" + stpnmb + "] " + temp2;
                                // 
                                populatedTr = populatedTr.Replace("$stepNo", temp);


                                // update comments
                                if (string.IsNullOrWhiteSpace(com) == false)
                                {
                                    populatedTr = populatedTr.Replace("$TC_Comments", com);
                                    populatedTr = populatedTr.Replace("$TC_Des_only", string.Empty);
                                }
                                // update step description
                                else
                                {
                                    populatedTr = populatedTr.Replace("$TC_Comments", string.Empty);
                                    populatedTr = populatedTr.Replace("$TC_Des_only", stpdesp);
                                }

                                string htmlAttach = string.Empty;
                                string htmlAttachPath = string.Empty;

                                // if Attachment (Screenshot) is present, Status should act as link
                                if (string.IsNullOrWhiteSpace(attch) == false)
                                {
                                    temp = "<a target='_blank' href='" + attch +
                                           "'><font style='color: $TestStepstatusColor; font-weight: " +
                                           Utility.GetParameter("ReportFailResultFontWeight") + ";'> " + stats +
                                           "</font></a>"; 
                                }
                                if (string.IsNullOrWhiteSpace(htmlattch) == false)
                                {
                                    htmlAttachPath = attch.Substring(0, attch.LastIndexOf("\\"));
                                    htmlAttach = "<a target='_blank' href='" + htmlattch +
                                                 "'><font style='color: $TestStepstatusColor; font-weight: normal;'> (html)</font></a>";
                                }
                                if (string.IsNullOrWhiteSpace(attch) && string.IsNullOrWhiteSpace(htmlattch))
                                {
                                    temp = "<font style='color: $TestStepstatusColor; font-weight: " +
                                           Utility.GetParameter("ReportPassResultFontWeight") + ";'>" + stats +
                                           "</font>";
                                }

                                // Update Status
                                populatedTr = populatedTr.Replace("$TC_Status", temp);

                                //view html link
                                populatedTr = populatedTr.Replace("$ViewHtml", htmlAttach);

                                // Update Execution Date and Time
                                temp = exedt + string.Empty + exetm;
                                
                                 populatedTr = populatedTr.Replace("$TC_ExecutionTime_Date", temp);

                                // Update Remarks
                                temp = rmrk;
                                populatedTr = populatedTr.Replace("$TC_Remarks", temp);

                                // Html Start + Expand/Collapse sign + individual TestCase Step (Initializing Cell Data)
                                finalHtml = finalHtml + populatedTr;

                                // updating each TestStep Status Tag Style
                                if (stats == ExecutionStatus.Fail)
                                {
                                    Common.Property.TotalStepFail++;

                                    // If Current TestCase Step is fail (means complete TestCase status will be fail)
                                    finalStatus = ExecutionStatus.Fail;
                                    finalStatusColor = failColor;
                                    // Html Start + Expand/Collapse sign + individual TestCase Step (Inintializing Cell Style)
                                    finalHtml = finalHtml.Replace("$TestStepstatusColor", failColor);

                                    //  Added remarks in test case header if test case has encountered a failed step.
                                    if (firstFail)
                                    {
                                        headerRemarkFail = rmrk;
                                        firstFail = false;
                                    }

                                    Property.FinalExecutionStatus = ExecutionStatus.Fail;
                                }
                                else if (stats == ExecutionStatus.Warning)
                                {
                                    Common.Property.TotalStepWarning++;

                                    if (finalStatus != ExecutionStatus.Fail)
                                        finalStatus = ExecutionStatus.Warning;
                                    if (finalStatusColor != failColor)
                                        finalStatusColor = warningColor;
                                    finalHtml = finalHtml.Replace("$TestStepstatusColor", warningColor);

                                    //  Added remarks in test case header if test case has encountered a warning step.
                                    if (firstWarning)
                                    {
                                        headerRemarkWarning = rmrk;
                                        firstWarning = false;
                                    }

                                }
                                else
                                {
                                    Common.Property.TotalStepPass++;

                                    // Html Start + Expand/Collapse sign + individual TestCase Step (Inintializing Cell Style)
                                    finalHtml = finalHtml.Replace("$TestStepstatusColor", passColor);
                                }

                                rowNo++;
                            }

                            #endregion

                            // Html Start + Expand/Collapse sign + individual TestCase Step (Status)
                            // Updating TestCase Status and style

                            finalHtml = finalHtml.Replace("$finalStatus", finalStatus);
                            //  set remarks in test case header
                            if (finalStatus == ExecutionStatus.Fail)
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_Remarks", headerRemarkFail);
                                Property.TotalCaseFail++;
                            }
                            else if (finalStatus == ExecutionStatus.Pass)
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_Remarks", "");
                                Property.TotalCasePass++;
                            }
                            else
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_Remarks", headerRemarkWarning);
                                Property.TotalCaseWarning++;
                            }

                            temp = "style='color: " + finalStatusColor + ";'";
                            // Html Start + Expand/Collapse sign + individual TestCase Step (Assigning Cell Style)
                            finalHtml = finalHtml.Replace("style='color: $FinalStatusColor;'", temp);
                            // Html Start + Expand/Collapse sign + individual TestCase Step (Updating Data Cell)
                            finalHtml = finalHtml + newTestCaseBodyEnd;

                            // Restoring the default value of finalStatus and finalStatusColor
                            finalStatus = ExecutionStatus.Pass;
                            finalStatusColor = passColor;



                            // Clearing the dataset
                            dsStore.Clear();
                            // Increasing the count for Expand/Collapse and TestCase Table ID
                            count++;

                            #endregion
                        }
                        catch
                        {
                            Console.WriteLine(Utility.GetCommonMsgVariable("KRYPTONERRCODE0006"));
                        }
                    }
                }

                finalHtml = finalHtml.Replace("$TotalCaseExecuted", Common.Property.TotalCaseExecuted.ToString());
                finalHtml = finalHtml.Replace("$TotalCasePass", Common.Property.TotalCasePass.ToString());
                finalHtml = finalHtml.Replace("$TotalCaseFail", Common.Property.TotalCaseFail.ToString());
                finalHtml = finalHtml.Replace("$TotalCaseWarning", Common.Property.TotalCaseWarning.ToString());


                // Ending Html Report
                finalHtml = finalHtml + htmlEnd;
                // Adding Complete Suit execution time




                TimeSpan time = executionEndTime - executionStartTime;


                finalHtml = finalHtml.Replace("$totalRunDuration", time.ToString());

                if (File.Exists(finalHtml)) File.Delete(finalHtml);

                StreamWriter sWriter = new StreamWriter(htmlFileLocation, false, Encoding.UTF8);
                // Saving Final Html Report
                sWriter.WriteLine(finalHtml);
                sWriter.Close();

                try
                {
                    if (File.Exists(companyLogo))
                    File.Copy(companyLogo, Common.Property.HtmlFileLocation + "\\" + generalLogoLocation);
                }
                catch
                {

                }
                //
                //compress all files into a zip file
                try
                {
                    using (ZipFile zip = new ZipFile())
                    {
                        foreach (string xmlLocation in xmlFilesLocation)
                        {
                            // Path to directory of files to compress and decompress.

                            string directory = Path.GetDirectoryName(xmlLocation);

                            FileInfo fileInfo = new FileInfo(xmlLocation);
                            zip.AddDirectory(directory, fileInfo.Directory.Parent.Name + "/" + fileInfo.Directory.Name);
                        }
                        zip.AddFile(Common.Property.HtmlFileLocation + "\\HtmlReport.html", "./");
                        if (File.Exists(Common.Property.HtmlFileLocation + "\\" + generalLogoLocation))
                        zip.AddFile(Common.Property.HtmlFileLocation + "\\" + generalLogoLocation, "./");
                        if (File.Exists(Common.Property.HtmlFileLocation + "\\" + Property.ReportZipFileName))//delete zip file if already exists
                            File.Delete(Common.Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);
                        zip.Save(Common.Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);


                    }
                }
                catch (Exception ex1)
                {
                    throw ex1;
                }

            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        ///  Creating Html Report from various XMLs
        /// </summary>
        /// <param></param>
        /// <returns></returns>    
        public static void CreateHtmlReport(string htmlFilename, bool zipRequired, bool logoUpload, bool sauceFlag)
        {
            try
            {
                string reportDestination = string.Empty;
                if (string.IsNullOrWhiteSpace(htmlFilename)) 
                { 
                    htmlFilename = "HtmlReport.html";
                    reportDestination = Common.Property.ResultsDestinationPath.ToString(); 
                }
                else 
                { 
                    reportDestination = Common.Property.HtmlFileLocation + "\\" + htmlFilename;
                }
                Common.Property.TIME = "HH:mm:ss";

                DateTime executionStartTime = DateTime.Now;
                DateTime executionEndTime = DateTime.Now;
                DateTime currentTime = executionEndTime;
                string generalLogoLocation = "logo" + Path.GetExtension(Property.CompanyLogo); 
                string passColor = Utility.GetParameter("ReportPassResultFontColor");
                string failColor = Utility.GetParameter("ReportFailResultFontColor");
                string warningColor = Utility.GetParameter("ReportWarningResultFontColor");
                string[] xmlFilesLocation = allXmlFilesLocation.Split(';');

                DataSet dsStore = new DataSet();
                string tcName = string.Empty;
                string tcID = string.Empty;
                string tcBrowser = string.Empty;
                string tcMachine = string.Empty;
                string tcStartTime = string.Empty;
                string tcEndTime = string.Empty;
                string tcDuration = string.Empty;

                string tcJobResultLocation = @"javascript:void(0)";

                string companyLogo = Property.CompanyLogo;
                string htmlFileLocation = Common.Property.HtmlFileLocation + "/" + htmlFilename;
                string finalStatus = ExecutionStatus.Pass;
                Property.FinalExecutionStatus = ExecutionStatus.Pass;
                string finalStatusColor = Utility.GetParameter("ReportPassResultFontColor");
                string populatedTr = string.Empty;
                string finalHtml = string.Empty;
                string temp = string.Empty;
                string tblID = string.Empty;
                string expander = string.Empty;
                string temp2 = string.Empty;
                string htmlStart =
                    @"<html>
                                    <head><style>td {overflow:hidden;}</style>
                                    <meta charset=UTF-8>
                                           <script>
                                                 function doMenu(item,obj) {
                                                                            tbl=item;
                                                                            if (tbl.style.display=='none') 
                                                                              {
                                                                                 tbl.style.display='block'; 
                                                                                 obj.innerHTML= '[-]';
                                                                               }
                                                                             else {
                                                                                  tbl.style.display='none';
                                                                                  obj.innerHTML= '[+]';
                                                                                  }
                                                                             }

                                            </script>
                                        </head>
                                            <body>
                                                 <div>                                                    
                                                    <p>
                                                    
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Execution Started at: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    @"$ExecutionStartTime</font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Execution Finished at: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    @"$ExecutionEndTime</font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Environment: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    Common.Utility.GetParameter(Common.Property.ENVIRONMENT) +
                    @"</font>
  <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Executed by: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    Property.KRYPTONVERSION +
                    @"</font> 
                   <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +" "+
                    @";'>  TestSuite:  </font>
                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>" +
                    Common.Utility.GetParameter("TestSuite").Trim() +
                      @"</font>
                        <img style='float: right;width:90px; float: right; margin-top: 6px;' src='" + generalLogoLocation + @"'/>
                         </p>
                                                    <p><font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Run Duration: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") +
                    @";'>$totalRunDuration</font>
                     </p>  
                    <p>                     
                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Test Case(s) Executed:</font>
                                                      <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @";  color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCaseExecuted</font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Passed:</font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @";  color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCasePass 
                    </font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Failed: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCaseFail 
                    </font>&nbsp;&nbsp;&nbsp;
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Total Warning: </font>
                                                    <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderValueFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderValueFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderValueFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderValueFontColor") + @";'>$TotalCaseWarning
                    </font>
                    </p>
                    
                                                <p>
                                                    </p></div><p>
                                              <font style='font-weight: " +
                    Utility.GetParameter("ReportHeaderTextFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportHeaderTextFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportHeaderTextFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportHeaderTextFontColor") +
                    @";'>Execution Report</font>
                                              <table width='100%' style='border:1px solid " +
                    Utility.GetParameter("ReportTableBorderColor") +
                    @"'>
                                                <td>
                                                    <table  id='header'  width='100%' style='table-layout:fixed'>
                                                         <tr style='font-weight: " +
                    Utility.GetParameter("ReportTableHeaderFontWeight") + @"; font-size: " +
                    Utility.GetParameter("ReportTableHeaderFontSize") + @"; font-style:" +
                    Utility.GetParameter("ReportTableHeaderFontStyle") + @"; font-family: " +
                    Utility.GetParameter("ReportFont") + @"; color: " +
                    Utility.GetParameter("ReportTableHeaderFontColor") + @"; background-color: " +
                    Utility.GetParameter("ReportTableHeaderBackgroundColor") +
                    @"'>
                                                         <td width='2%'></td>
                                                         <td width='35%'>Steps</td>                                                         
                                                         <td width='8%'>Status</td>
                                                        <td width='5%'>Browser</td>
                                                         <td width='10%'>Machine</td>
                                                         <td width='15%'>Execution Time</td>
                                                         <td width='25%'>Remarks</td>                                                            
                                                    </table>";


                // added remarks in each test case header
                //                     Added date/time of execution in each test case header
                string newTestCaseBodyStart = @"<table width='100%' style='table-layout:fixed'>
                                            <tr style='font-family: " + Utility.GetParameter("ReportFont") + @"; background-color: " + Utility.GetParameter("ReportTestCaseBackgroundColor") + @"'>
                                             <td width='2%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'><a id='$xn' href='javascript:doMenu($TblID,$xn)'>[+]</a></td>
                                             <td width='35%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>[$TCId] $TCName</td>                                             
                                             <td width='8%' style='font-weight: " + Utility.GetParameter("ReportTestCaseStatusWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; '><span id='span_1' style='color: $FinalStatusColor;'>$finalStatus</span></td>
                                             <td width='5%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_Browser</td>
                                            <td width='10%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_Machine</td>
                                             <td width='15%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_ExecutionTime_Date</td>
                                             <td width='25%' style='font-weight: " + Utility.GetParameter("ReportRemarksFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportRemarksFontStyle") + @"; color: " + Utility.GetParameter("ReportRemarksFontColor") + @"'>$Header_TC_Remarks</td>                                            
                                            </tr>
                                           </table>
                                           <div id='$TblID' style='display:none'>";

                //forcing to always use this header
                if ((sauceFlag) || (true))
                {
                    newTestCaseBodyStart = @"<table width='100%' style='table-layout:fixed'>
                                            <tr style='font-family: " + Utility.GetParameter("ReportFont") + @"; background-color: " + Utility.GetParameter("ReportTestCaseBackgroundColor") + @"'>
                                             <td width='2%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'><a id='$xn' href='javascript:doMenu($TblID,$xn)'>[+]</a></td>
                                             <td width='35%' title='$TCId' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'><b>$TCName</b></td>                                             
                                             <td width='8%' style='font-weight: " + Utility.GetParameter("ReportTestCaseStatusWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; '><span id='span_1' style='color: $FinalStatusColor;'><a target='_blank' href='" + "$TCJobLocation" + @"' style='text-decoration: none'><font style='color: $TestHeadertatusColor; font-weight: " + Utility.GetParameter("ReportFailResultFontWeight") + @";'> $finalStatus  </font></a></span></td>
                                             <td width='5%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_Browser</td>
                                            <td width='10%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_Machine</td>
                                             <td width='15%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$Header_TC_ExecutionTime_Date</td>
                                             <td width='25%' style='font-weight: " + Utility.GetParameter("ReportRemarksFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportRemarksFontStyle") + @"; color: " + Utility.GetParameter("ReportRemarksFontColor") + @"'>$Header_TC_Remarks</td>                                            
                                            </tr>
                                           </table>
                                           <div id='$TblID' style='display:none'>";
                }
                // Introduction of one more column (only Description)-
                string trStructure = @"<table width='100%' style='word-wrap: break-word; table-layout: fixed;'> 
                                        <tr style='font-family: " + Utility.GetParameter("ReportFont") + @"; background-color: $trBackground;'>
                                            <td width='2%' ></td>
                                            <td width='35%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$stepNo<!--<font style='font-style:italic'>-->$TC_Comments<!--</font>-->$TC_Des_only</td>                                             
                                            <td width='8%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; '>$TC_Status $ViewHtml</td> 
                                            <td width='5%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>&nbsp;</td>
                                            <td width='10%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>&nbsp;</td>    
                                            <td width='15%' style='font-weight: " + Utility.GetParameter("ReportFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportFontStyle") + @"; color: " + Utility.GetParameter("ReportFontColor") + @"'>$TC_ExecutionTime_Date</td>
                                            <td width='25%' style='font-weight: " + Utility.GetParameter("ReportRemarksFontWeight") + @"; font-size: " + Utility.GetParameter("ReportFontSize") + @"; font-style:" + Utility.GetParameter("ReportRemarksFontStyle") + @"; color: " + Utility.GetParameter("ReportRemarksFontColor") + @"'>$TC_Remarks</td>                                                                                           
                                        </tr>
                                     </table>";
                string newTestCaseBodyEnd = @"</div>";

                string htmlEnd = @"</td>
                                 </tr>
                                  </table>
                                            </body>
                                </html>";

                //

                // Html Start
                finalHtml = htmlStart;

                int count = 0;


                foreach (string xmlLocation in xmlFilesLocation)
                {
                    if (File.Exists(xmlLocation))
                    {
                        string rootFolderName = new FileInfo(xmlLocation).Directory.Parent.Name;
                        string rootFolderExt = string.Empty;
                        try
                        {
                            if (rootFolderName.Contains('-'))
                            rootFolderExt = rootFolderName.Substring(rootFolderName.IndexOf('-'),
                                                                            rootFolderName.Length -
                                                                            rootFolderName.IndexOf('-'));
                        }
                        catch
                        {

                        }
                        Property.TotalCaseExecuted++;

                        try
                        {
                            #region Individual TestCase

                            dsStore = new DataSet();
                            // copying all data from XML to datatable
                            dsStore.ReadXml(xmlLocation);
                            DataTable dtTemp = new DataTable();

                            dtTemp = dsStore.Tables[0];

                            // getting TestCase Name and ID from extended properties
                            if (dtTemp.ExtendedProperties["TestCase Name"] != null)
                                tcName = dtTemp.ExtendedProperties["TestCase Name"].ToString();
                            if (dtTemp.ExtendedProperties["TestCase Id"] != null)
                                tcID = dtTemp.ExtendedProperties["TestCase Id"].ToString();
                            if (dtTemp.ExtendedProperties["Browser"] != null)
                                tcBrowser = dtTemp.ExtendedProperties["Browser"].ToString();
                            if (dtTemp.ExtendedProperties["RCMachineId"] != null)
                                tcMachine = dtTemp.ExtendedProperties["RCMachineId"].ToString();
                            if ((dtTemp.ExtendedProperties["SauceJobUrl"] != null) && (dtTemp.ExtendedProperties["SauceJobUrl"].ToString() != ""))
                                tcJobResultLocation = dtTemp.ExtendedProperties["SauceJobUrl"].ToString();

                            //Extract date/time information from xml files
                            if (dtTemp.ExtendedProperties["ExecutionEndTime"] != null)
                                tcEndTime = dtTemp.ExtendedProperties["ExecutionEndTime"].ToString();
                            if (dtTemp.ExtendedProperties["ExecutionStartTime"] != null)
                                tcStartTime = dtTemp.ExtendedProperties["ExecutionStartTime"].ToString();
                            if (dtTemp.ExtendedProperties["ExecutionDuration"] != null)
                                tcDuration = dtTemp.ExtendedProperties["ExecutionDuration"].ToString();



                            #region Updating calculation of start and end date time, section below commented
                            try
                            {
                                DateTime xmlStartTime = DateTime.ParseExact(tcStartTime.ToString(), Property.DATE_TIME, null);
                                DateTime xmlEndTime = DateTime.ParseExact(tcEndTime.ToString(), Property.DATE_TIME, null);
                                //Set start time
                                if (executionStartTime > xmlStartTime)
                                    executionStartTime = xmlStartTime;

                                //Set end time
                                if ((xmlEndTime > executionEndTime) || (executionEndTime.Equals(currentTime)))
                                    executionEndTime = xmlEndTime;
                            }
                            catch
                            {
                            }
                            #endregion

                            tblID = "tblID_" + count;
                            expander = "x_" + count;

                            temp = newTestCaseBodyStart;
                            temp = temp.Replace("$TCName", tcName);
                            temp = temp.Replace("$TCId", tcID);
                            temp = temp.Replace("$TblID", tblID);
                            temp = temp.Replace("$xn", expander);
                            // Html Start + Expand/Collapse sign (for each TestCase)
                            finalHtml = finalHtml + temp;

                           

                            #region Each TestCase Step

                            int rowNo = 0;

                            //  flag to keep track of first warning/failed step.
                            bool firstFail = true;
                            bool firstWarning = true;
                            //  headerRemark is empty if test case is passed. Else it will be updated on a failed or warning step.
                            string headerRemarkFail = string.Empty;
                            string headerRemarkWarning = string.Empty;

                            //  added start date/time in test case header
                            if (tcDuration == string.Empty)
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_ExecutionTime_Date", dtTemp.Rows[0][3].ToString() + " : " + dtTemp.Rows[0][4].ToString());
                            }
                            else
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_ExecutionTime_Date", tcDuration);
                            }

                            //update browser info
                            finalHtml = finalHtml.Replace("$Header_TC_Browser", tcBrowser);

                            //update machine info
                            finalHtml = finalHtml.Replace("$Header_TC_Machine", tcMachine);

                            //Update sauce job url
                            finalHtml = finalHtml.Replace("$TCJobLocation", tcJobResultLocation);




                            foreach (DataRow tcRow in dtTemp.Rows) // Loop over the rows.
                            {
                                Common.Property.TotalStepExecuted++;

                                // taking basic TR structure
                                populatedTr = trStructure;

                                // populating row with data (present in Xml row)
                                string stpnmb = tcRow.ItemArray[0].ToString();
                                string stpdesp = tcRow.ItemArray[1].ToString();
                                string stats = tcRow.ItemArray[2].ToString();
                                string exedt = tcRow.ItemArray[3].ToString();
                                string exetm = tcRow.ItemArray[4].ToString();
                                string rmrk = tcRow.ItemArray[5].ToString();
                                string attch = tcRow.ItemArray[6].ToString();
                                string htmlattch = tcRow.ItemArray[7].ToString();
                                string objhigh = tcRow.ItemArray[8].ToString();
                                string com = tcRow.ItemArray[9].ToString();
                                string attachurl = tcRow.ItemArray[10].ToString();

                                if (string.IsNullOrWhiteSpace(attch) == false)
                                {
                                    try
                                    {
                                        attch = attch.Substring(0, attch.IndexOf('\\')) + rootFolderExt +
                                                attch.Substring(attch.IndexOf('\\'), attch.Length - attch.IndexOf('\\'));
                                    }
                                    catch { }
                                }
                                if (string.IsNullOrWhiteSpace(htmlattch) == false)
                                {
                                    try
                                    {
                                        htmlattch = htmlattch.Substring(0, htmlattch.IndexOf('\\')) + rootFolderExt +
                                                    htmlattch.Substring(htmlattch.IndexOf('\\'),
                                                                        htmlattch.Length - htmlattch.IndexOf('\\'));
                                    }
                                    catch { }
                                }





                                int outParam = 0;
                                Math.DivRem(rowNo, 2, out outParam);
                                if (outParam > 0)
                                {
                                    populatedTr = populatedTr.Replace("$trBackground",
                                                                      Utility.GetParameter(
                                                                          "ReportTableAlternateRowColor"));
                                }
                                else
                                {
                                    populatedTr = populatedTr.Replace("$trBackground",
                                                                      Utility.GetParameter("ReportTableRowColor"));
                                }



                                temp = "[" + stpnmb + "] " + temp2;
                                // 
                                populatedTr = populatedTr.Replace("$stepNo", temp);


                                // update comments
                                if (string.IsNullOrWhiteSpace(com) == false)
                                {
                                    populatedTr = populatedTr.Replace("$TC_Comments", com);
                                    populatedTr = populatedTr.Replace("$TC_Des_only", string.Empty);
                                }
                                // update step description
                                else
                                {
                                    populatedTr = populatedTr.Replace("$TC_Comments", string.Empty);
                                    populatedTr = populatedTr.Replace("$TC_Des_only", stpdesp);
                                }

                                string htmlAttach = string.Empty;
                                string htmlAttachPath = string.Empty;

                                // if Attachment (Screenshot) is present, Status should act as link
                                if (string.IsNullOrWhiteSpace(attch) == false)
                                {
                                    temp = "<a target='_blank' title='" + attachurl + "' href='" + attch +
                                           "'><font style='color: $TestStepstatusColor; font-weight: " +
                                           Utility.GetParameter("ReportFailResultFontWeight") + ";'> " + stats +
                                           "</font></a>"; 
                                }
                                try
                                {
                                    if (string.IsNullOrWhiteSpace(htmlattch) == false)
                                    {
                                       
                                        if (string.IsNullOrWhiteSpace(attch))
                                        {
                                            temp = "<a target='_blank' title='" + attachurl + "'><font style='color: $TestStepstatusColor; font-weight: " +
                                             Utility.GetParameter("ReportFailResultFontWeight") + ";'> " + stats + "<a target='_blank' href='" + htmlattch +"'><font style='color: $TestStepstatusColor; font-weight: normal;'> (html)</font></a>"; // Added By : To attach html report with the status properly in Case of HTML and IMAGE only. 
                                        }
                                        else
                                        {
                                            htmlAttach = "<a target='_blank' href='" + htmlattch +
                                                        "'><font style='color: $TestStepstatusColor; font-weight: normal;'> (html)</font></a>";
                                        }
                                    }

                                }
                                catch { }
                                if (string.IsNullOrWhiteSpace(attch) && string.IsNullOrWhiteSpace(htmlattch))
                                {
                                    temp = "<font style='color: $TestStepstatusColor; font-weight: " +
                                           Utility.GetParameter("ReportPassResultFontWeight") + ";'>" + stats +
                                           "</font>";
                                }

                                // Update Status
                                populatedTr = populatedTr.Replace("$TC_Status", temp);

                                //view html link
                                populatedTr = populatedTr.Replace("$ViewHtml", htmlAttach);

                                // Update Execution Date and Time
                                temp = exedt + string.Empty + exetm;
                                populatedTr = populatedTr.Replace("$TC_ExecutionTime_Date", temp);

                                // Update Remarks
                                temp = rmrk;
                                populatedTr = populatedTr.Replace("$TC_Remarks", temp);

                                finalHtml = finalHtml + populatedTr;

                                // updating each TestStep Status Tag Style
                                if (stats == ExecutionStatus.Fail)
                                {
                                    Common.Property.TotalStepFail++;

                                    // If Current TestCase Step is fail (means complete TestCase status will be fail)
                                    finalStatus = ExecutionStatus.Fail;
                                    finalStatusColor = failColor;
                                    
                                    finalHtml = finalHtml.Replace("$TestStepstatusColor", failColor);

                                    //  Added remarks in test case header if test case has encountered a failed step.
                                    if (firstFail)
                                    {
                                        headerRemarkFail = rmrk;
                                        firstFail = false;
                                    }

                                    Property.FinalExecutionStatus = ExecutionStatus.Fail;
                                }
                                else if (stats == ExecutionStatus.Warning)
                                {
                                    Common.Property.TotalStepWarning++;

                                    if (finalStatus != ExecutionStatus.Fail)
                                        finalStatus = ExecutionStatus.Warning;
                                    if (finalStatusColor != failColor)
                                        finalStatusColor = warningColor;
                                    finalHtml = finalHtml.Replace("$TestStepstatusColor", warningColor);

                                    //  Added remarks in test case header if test case has encountered a warning step.
                                    if (firstWarning)
                                    {
                                        headerRemarkWarning = rmrk;
                                        firstWarning = false;
                                    }

                                }
                                else
                                {
                                    Common.Property.TotalStepPass++;

                                    
                                    finalHtml = finalHtml.Replace("$TestStepstatusColor", passColor);
                                }

                                rowNo++;
                            }

                            #endregion

                            // Updating TestCase Status and style

                            finalHtml = finalHtml.Replace("$finalStatus", finalStatus);
                            if (finalStatus.Equals(ExecutionStatus.Pass)) { finalHtml = finalHtml.Replace("$TestHeadertatusColor", passColor); }
                            else if (finalStatus.Equals(ExecutionStatus.Fail)) { finalHtml = finalHtml.Replace("$TestHeadertatusColor", failColor); }
                            else if (finalStatus.Equals(ExecutionStatus.Warning)) { finalHtml = finalHtml.Replace("$TestHeadertatusColor", warningColor); }
                            //  set remarks in test case header
                            if (finalStatus == ExecutionStatus.Fail)
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_Remarks", headerRemarkFail);
                                Property.TotalCaseFail++;
                                Property.ExecutionFailReason = headerRemarkFail;
                            }
                            else if (finalStatus == ExecutionStatus.Pass)
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_Remarks", string.Empty);
                                Property.TotalCasePass++;
                                
                            }
                            else
                            {
                                finalHtml = finalHtml.Replace("$Header_TC_Remarks", headerRemarkWarning);
                                Property.TotalCaseWarning++;

                            }
                            temp = "style='color: " + finalStatusColor + ";'";
                            
                            finalHtml = finalHtml.Replace("style='color: $FinalStatusColor;'", temp);
                            
                            finalHtml = finalHtml + newTestCaseBodyEnd;

                            // Restoring the default value of finalStatus and finalStatusColor
                            finalStatus = ExecutionStatus.Pass;
                            finalStatusColor = passColor;


                            // Clearing the dataset
                            dsStore.Clear();
                            // Increasing the count for Expand/Collapse and TestCase Table ID
                            count++;

                            #endregion
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                }

                finalHtml = finalHtml.Replace("$TotalCaseExecuted", Common.Property.TotalCaseExecuted.ToString());
                finalHtml = finalHtml.Replace("$TotalCasePass", Common.Property.TotalCasePass.ToString());
                finalHtml = finalHtml.Replace("$TotalCaseFail", Common.Property.TotalCaseFail.ToString());
                finalHtml = finalHtml.Replace("$TotalCaseWarning", Common.Property.TotalCaseWarning.ToString());
                // Ending Html Report
                finalHtml = finalHtml + htmlEnd;
                // Adding Complete Suit execution time

                finalHtml = finalHtml.Replace("$ExecutionStartTime", executionStartTime.ToString(Utility.GetParameter("DateTimeFormat")));
                finalHtml = finalHtml.Replace("$ExecutionEndTime", executionEndTime.ToString(Utility.GetParameter("DateTimeFormat")));

                TimeSpan time = executionEndTime - executionStartTime;

                Common.Property.ExecutionStartDateTime = executionStartTime.ToString(Utility.GetParameter("DateTimeFormat"));
                Common.Property.ExecutionEndDateTime = executionEndTime.ToString(Utility.GetParameter("DateTimeFormat"));
                Common.Utility.SetVariable("ExecutionStartDateTime", Common.Property.ExecutionStartDateTime);
                Common.Utility.SetVariable("ExecutionEndDateTime", Common.Property.ExecutionEndDateTime);


                finalHtml = finalHtml.Replace("$totalRunDuration", time.ToString());

                Common.Property.ReportSummaryBody = finalHtml;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }



    }

}
