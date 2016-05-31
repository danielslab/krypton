/****************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Reporting.LogFile.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Creating Xml and Html Reports
** **************************************************************************/

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Data;
using Common;
using Ionic.Zip;


namespace Reporting
{

    public sealed class LogFile
    {
        // datatable by which Xml will be created
         private static DataTable _testCaseTable;
         private static LogFile _logInstance = null;
        // string which contains Xml file path
        public string ExecutedXmlPath;
        public static string AllXmlFilesLocation = string.Empty;
        public static TimeSpan TimeTaken = new TimeSpan(0, 0, 0);
        public static Boolean  FirstEntryInLogFile=true;
        public string ExecutionTime;

        private LogFile(string filePathName)
        {
            _testCaseTable = new DataTable("TestStep");
            // Copy the file name -- it will be used in function executeStep
            ExecutedXmlPath = filePathName;
            // Storing all Xml Path file names at one place (will be used while generating HTML report)
            AllXmlFilesLocation = AllXmlFilesLocation + "," + filePathName;

            // Rows Structure
            _testCaseTable.Columns.Add("StepNumber", typeof(string));
            _testCaseTable.Columns.Add("StepDescription", typeof(string));
            _testCaseTable.Columns.Add("Status", typeof(string));
            _testCaseTable.Columns.Add("ExecutionDate", typeof(string));
            _testCaseTable.Columns.Add("ExecutionTime", typeof(string));
            _testCaseTable.Columns.Add("Remarks", typeof(string));
            _testCaseTable.Columns.Add("Attachments", typeof(string));
            _testCaseTable.Columns.Add("HtmlAttachments", typeof(string));
            _testCaseTable.Columns.Add("ObjectHighlight", typeof(string));
            _testCaseTable.Columns.Add("StepComments", typeof(string));
            _testCaseTable.Columns.Add("AttachmentUrl", typeof(string));


            //creating the folder structure here
            string directory = Path.GetDirectoryName(filePathName);
            if (directory != null) CreateDirectory(new DirectoryInfo(directory));
        }

        #region Properties
        public static string filePathName { get; set; }

        public static LogFile Instance
        {
            get
            {
                _logInstance = new LogFile(filePathName);
                return _logInstance;
            }
        }

        
        #endregion


        /// <summary>
        ///Creating folders and sub-folders. It is called iteratively 
        /// till a already present directory is not found.
        /// </summary>
        /// <param name="directory">This will contain Folder name and if not present 
        /// then called again and in end sub child folders are formed one by one</param>
        /// <returns></returns>
        private static void CreateDirectory(DirectoryInfo directory)
        {
            if (directory.Parent != null && !directory.Parent.Exists)
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
            string stepNumber = Property.StepNumber;
            string stepDescription = Property.StepDescription;
           string status = Property.Status;
            //Updated Code to display accurate Time in seconds and milliseconds of each stepAction 
            if (FirstEntryInLogFile)
            {
                ExecutionTime = String.Format("{0:00.00}", (DateTime.Now.TimeOfDay - Convert.ToDateTime(Property.ExecutionTime).TimeOfDay).TotalSeconds); //only first time in log file
                FirstEntryInLogFile = false;
            }
            else
            {
                ExecutionTime = String.Format("{0:00.00}", (DateTime.Now.TimeOfDay - TimeTaken).TotalSeconds);

            }
             TimeTaken = DateTime.Now.TimeOfDay;            
         
            string remarks = Property.Remarks;
            string attachments = Property.Attachments;
            string attachmentUrl = Property.AttachmentsUrl;
            string htmlAttachments = Property.HtmlSourceAttachment;
            string objectHighlight = Property.ObjectHighlight;
            string stepComments = Property.StepComments;
            if (status.Equals(ExecutionStatus.Fail))
                Property.JobExecutionStatus = status;

            // Creating a new row for table and pupulating row's cells
            DataRow dr = _testCaseTable.NewRow();
            dr["StepNumber"] = stepNumber;
            dr["StepDescription"] = stepDescription;
            dr["Status"] = status;
            dr["ExecutionTime"] = ExecutionTime;           
            dr["Remarks"] = remarks;
            if (attachments != string.Empty) // when some image is present
            {
                DirectoryInfo drInfo = new DirectoryInfo(ExecutedXmlPath);
                if (drInfo.Parent != null)
                {
                    if (drInfo.Parent.Parent != null)
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
                DirectoryInfo drInfo = new DirectoryInfo(ExecutedXmlPath);
                if (drInfo.Parent != null)
                {
                    if (drInfo.Parent.Parent != null)
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
                Property.JobExecutionStatus = ExecutionStatus.Fail;
                failCounter++;
                if (failCounter.Equals(1))
                    Property.ExecutionFailReason = remarks;
            }
            // adding newly created row to table
            _testCaseTable.Rows.Add(dr);

        }

        public void SaveXmlLog()
        {
            //creating a new Xml file (replacing old Xml with new one)
            _testCaseTable.WriteXml(ExecutedXmlPath, XmlWriteMode.WriteSchema);
        }

        ///  <summary>
        /// For adding new attributes to datatable (used in creating Xml Report)
        ///  </summary>
        ///  <param></param>
        /// <param name="attribute"></param>
        /// <param name="value"></param>
        /// <returns></returns> 
        public void AddTestAttribute(string attribute, string value)
        {
            _testCaseTable.ExtendedProperties[attribute] = value;

        }

        /// <summary>
        ///Creating Html Report from various XMLs
        /// </summary>
        /// <param></param>
        /// <returns></returns>    
        public void CreateHtmlReport1()
        {
            DateTime executionStartTime = DateTime.ParseExact(Property.ExecutionStartDateTime, Utility.GetParameter("DateTimeFormat"),
                null);
            DateTime executionEndTime = DateTime.ParseExact(Property.ExecutionEndDateTime, Utility.GetParameter("DateTimeFormat"),
                null);
            string generalLogoLocation = "logo" + Path.GetExtension(Property.CompanyLogo); 
            string passColor = Utility.GetParameter("ReportPassResultFontColor");
            string failColor = Utility.GetParameter("ReportFailResultFontColor");
            string warningColor = Utility.GetParameter("ReportWarningResultFontColor");
            string browserDetails = Utility.GetParameter(Property.BrowserString); 
            string browserVersion = Utility.GetVariable(Property.BrowserVersion);
            string[] xmlFilesLocation = AllXmlFilesLocation.Split(';');
            string tcName = string.Empty;
            string tcId = string.Empty;
            string tcBrowser = string.Empty;

            string companyLogo = Property.CompanyLogo;
            string htmlFileLocation = Property.HtmlFileLocation + "/HtmlReport.html";
            string finalStatus = ExecutionStatus.Pass;
            Property.FinalExecutionStatus = ExecutionStatus.Pass;
            string finalStatusColor = Utility.GetParameter("ReportPassResultFontColor");
            StringBuilder finalHtml = new StringBuilder();
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
                executionStartTime.ToString(Utility.GetParameter("DateTimeFormat")) +
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
                executionEndTime.ToString(Utility.GetParameter("DateTimeFormat")) +
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
                Utility.GetParameter(Property.Environment) + 
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
            finalHtml.Append(htmlStart);

            int count = 0;


            foreach (string xmlLocation in xmlFilesLocation)
            {
                if (File.Exists(xmlLocation))
                {
                    Property.TotalCaseExecuted++;

                    try
                    {
                        #region Individual TestCase

                        var dsStore = new DataSet();
                        // copying all data from XML to datatable
                        dsStore.ReadXml(xmlLocation);

                        var dtTemp = dsStore.Tables[0];

                        // getting TestCase Name and ID from extended properties
                        if (dtTemp.ExtendedProperties["TestCase Name"] != null)
                            tcName = dtTemp.ExtendedProperties["TestCase Name"].ToString(); 
                        if (dtTemp.ExtendedProperties["TestCase Id"] != null)
                            tcId = dtTemp.ExtendedProperties["TestCase Id"].ToString(); 
                        if (dtTemp.ExtendedProperties["Browser"] != null)
                            tcBrowser = dtTemp.ExtendedProperties["Browser"].ToString(); 

                        var tblId = "tblID_" + count;
                        var expander = "x_" + count;

                        var temp = newTestCaseBodyStart;
                        temp = temp.Replace("$TCName", tcName);
                        temp = temp.Replace("$TCId", tcId);
                        temp = temp.Replace("$TblID", tblId);
                        temp = temp.Replace("$xn", expander);
                        // Html Start + Expand/Collapse sign (for each TestCase)
                        finalHtml.Append(temp);

                        // Calculating time duration for each TestCase

                        #region Each TestCase Step

                        int rowNo = 0;

                        //flag to keep track of first warning/failed step.
                        bool firstFail = true;
                        bool firstWarning = true;
                        //headerRemark is empty if test case is passed. Else it will be updated on a failed or warning step.
                        string headerRemarkFail = string.Empty;
                        string headerRemarkWarning = string.Empty;

                        //added start date/time in test case header
                        finalHtml = finalHtml.Replace("$Header_TC_ExecutionTime_Date", dtTemp.Rows[0][3] + " : " + dtTemp.Rows[0][4]);

                        //update browser info
                        finalHtml = finalHtml.Replace("$Header_TC_Browser", tcBrowser);


                        foreach (DataRow tcRow in dtTemp.Rows) // Loop over the rows.
                        {
                            Property.TotalStepExecuted++;

                            // taking basic TR structure
                            var populatedTr = trStructure;

                            // populating row with data (present in Xml row)
                            string stpnmb = tcRow.ItemArray[0].ToString();
                            string stpdesp = tcRow.ItemArray[1].ToString();
                            string stats = tcRow.ItemArray[2].ToString();
                            string exedt = tcRow.ItemArray[3].ToString();
                            string exetm = tcRow.ItemArray[4].ToString();
                            string rmrk = tcRow.ItemArray[5].ToString();
                            string attch = tcRow.ItemArray[6].ToString();
                            string htmlattch = tcRow.ItemArray[7].ToString();
                            string com = tcRow.ItemArray[9].ToString();
                            int outParam;
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
                            finalHtml.Append(populatedTr);

                            // updating each TestStep Status Tag Style
                            if (stats == ExecutionStatus.Fail)
                            {
                                Property.TotalStepFail++;

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
                                Property.TotalStepWarning++;

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
                                Property.TotalStepPass++;
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
                            finalHtml = finalHtml.Replace("$Header_TC_Remarks", string.Empty);
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
                        finalHtml.Append(newTestCaseBodyEnd);

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

            finalHtml = finalHtml.Replace("$TotalCaseExecuted", Property.TotalCaseExecuted.ToString());
            finalHtml = finalHtml.Replace("$TotalCasePass", Property.TotalCasePass.ToString());
            finalHtml = finalHtml.Replace("$TotalCaseFail", Property.TotalCaseFail.ToString());
            finalHtml = finalHtml.Replace("$TotalCaseWarning", Property.TotalCaseWarning.ToString());
            // Ending Html Report
            finalHtml.Append( htmlEnd);
            // Adding Complete Suit execution time
            TimeSpan time = executionEndTime - executionStartTime;
            finalHtml = finalHtml.Replace("$totalRunDuration", time.ToString());

            if (File.Exists(Convert.ToString(finalHtml))) File.Delete(Convert.ToString(finalHtml));

            StreamWriter sWriter = new StreamWriter(htmlFileLocation, false, Encoding.UTF8);
            // Saving Final Html Report
            sWriter.WriteLine(finalHtml);
            sWriter.Close();

            try
            {
                if (File.Exists(companyLogo))
                    File.Copy(companyLogo, Property.HtmlFileLocation + "\\" + generalLogoLocation);
            }
            catch
            {
                // ignored
            }
            //
            //compress all files into a zip file
            using (ZipFile zip = new ZipFile())
            {
                foreach (string xmlLocation in xmlFilesLocation)
                {
                    // Path to directory of files to compress and decompress.

                    string directory = Path.GetDirectoryName(xmlLocation);

                    FileInfo fileInfo = new FileInfo(xmlLocation);
                    zip.AddDirectory(directory, fileInfo.Directory.Parent.Name + "/" + fileInfo.Directory.Name);
                }
                zip.AddFile(Property.HtmlFileLocation + "\\HtmlReport.html", "./");
                if (File.Exists(Property.HtmlFileLocation + "\\" + generalLogoLocation))
                    zip.AddFile(Property.HtmlFileLocation + "\\" + generalLogoLocation, "./");
                if (File.Exists(Property.HtmlFileLocation + "\\" + Property.ReportZipFileName))//delete zip file if already exists
                    File.Delete(Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);
                zip.Save(Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);


            }
        }

        /// <summary>
        ///  Creating Html Report from various XMLs
        /// </summary>
        /// <param></param>
        /// <param name="htmlFilename"></param>
        /// <param name="zipRequired"></param>
        /// <param name="logoUpload"></param>
        /// <param name="sauceFlag"></param>
        /// <returns></returns>    
        public static void CreateHtmlReport(string htmlFilename, bool zipRequired, bool logoUpload, bool sauceFlag)
        {
            try
            {
                Property.Time = "HH:mm:ss";
                DateTime executionStartTime = DateTime.Now;
                DateTime executionEndTime = DateTime.Now;
                DateTime currentTime = executionEndTime;
                string generalLogoLocation = "logo" + Path.GetExtension(Property.CompanyLogo); 
                string passColor = Utility.GetParameter("ReportPassResultFontColor");
                string failColor = Utility.GetParameter("ReportFailResultFontColor");
                string warningColor = Utility.GetParameter("ReportWarningResultFontColor");
                string[] xmlFilesLocation = AllXmlFilesLocation.Split(';');

                string tcName = string.Empty;
                string tcId = string.Empty;
                string tcBrowser = string.Empty;
                string tcMachine = string.Empty;
                string tcStartTime = string.Empty;
                string tcEndTime = string.Empty;
                string tcDuration = string.Empty;
                string tcJobResultLocation = @"javascript:void(0)";
                string finalStatus = ExecutionStatus.Pass;
                Property.FinalExecutionStatus = ExecutionStatus.Pass;
                string finalStatusColor = Utility.GetParameter("ReportPassResultFontColor");
                StringBuilder finalHtml = new StringBuilder();
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
                    Utility.GetParameter(Property.Environment) +
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
                    Property.KryptonVersion +
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
                    Utility.GetParameter("TestSuite").Trim() +
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

                //forcing to always use this header
                var newTestCaseBodyStart = @"<table width='100%' style='table-layout:fixed'>
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
                finalHtml.Append(htmlStart);

                int count = 0;


                foreach (string xmlLocation in xmlFilesLocation)
                {
                    //using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
                    //{
                    //    sw.WriteLine("Logfile Foreach 1");
                    //}
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
                            // ignored
                        }
                        Property.TotalCaseExecuted++;

                        try
                        {
                            #region Individual TestCase

                            var dsStore = new DataSet();
                            // copying all data from XML to datatable
                            dsStore.ReadXml(xmlLocation);

                            var dtTemp = dsStore.Tables[0];

                            // getting TestCase Name and ID from extended properties
                            if (dtTemp.ExtendedProperties["TestCase Name"] != null)
                                tcName = dtTemp.ExtendedProperties["TestCase Name"].ToString();
                            if (dtTemp.ExtendedProperties["TestCase Id"] != null)
                                tcId = dtTemp.ExtendedProperties["TestCase Id"].ToString();
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
                                DateTime xmlStartTime = DateTime.ParseExact(tcStartTime, Property.Date_Time, null);
                                DateTime xmlEndTime = DateTime.ParseExact(tcEndTime, Property.Date_Time, null);
                                //Set start time
                                if (executionStartTime > xmlStartTime)
                                    executionStartTime = xmlStartTime;

                                //Set end time
                                if ((xmlEndTime > executionEndTime) || (executionEndTime.Equals(currentTime)))
                                    executionEndTime = xmlEndTime;
                            }
                            catch
                            {
                                // ignored
                            }

                            #endregion

                            var tblId = "tblID_" + count;
                            var expander = "x_" + count;

                            var temp = newTestCaseBodyStart;
                            temp = temp.Replace("$TCName", tcName);
                            temp = temp.Replace("$TCId", tcId);
                            temp = temp.Replace("$TblID", tblId);
                            temp = temp.Replace("$xn", expander);
                            // Html Start + Expand/Collapse sign (for each TestCase)
                            finalHtml.Append(temp);

                           

                            #region Each TestCase Step

                            int rowNo = 0;

                            //  flag to keep track of first warning/failed step.
                            bool firstFail = true;
                            bool firstWarning = true;
                            //  headerRemark is empty if test case is passed. Else it will be updated on a failed or warning step.
                            string headerRemarkFail = string.Empty;
                            string headerRemarkWarning = string.Empty;

                            //  added start date/time in test case header
                            finalHtml = tcDuration == string.Empty ? finalHtml.Replace("$Header_TC_ExecutionTime_Date", dtTemp.Rows[0][3] + " : " + dtTemp.Rows[0][4]) : finalHtml.Replace("$Header_TC_ExecutionTime_Date", tcDuration);

                            //update browser info
                            finalHtml = finalHtml.Replace("$Header_TC_Browser", tcBrowser);

                            //update machine info
                            finalHtml = finalHtml.Replace("$Header_TC_Machine", tcMachine);

                            //Update sauce job url
                            finalHtml = finalHtml.Replace("$TCJobLocation", tcJobResultLocation);




                            foreach (DataRow tcRow in dtTemp.Rows) // Loop over the rows.
                            {
                                //using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
                                //{
                                //    sw.WriteLine("Logfile Foreach 2");
                                //}
                                Property.TotalStepExecuted++;

                                // taking basic TR structure
                                var populatedTr = trStructure;

                                // populating row with data (present in Xml row)
                                string stpnmb = tcRow.ItemArray[0].ToString();
                                string stpdesp = tcRow.ItemArray[1].ToString();
                                string stats = tcRow.ItemArray[2].ToString();
                                string exedt = tcRow.ItemArray[3].ToString();
                                string exetm = tcRow.ItemArray[4].ToString();
                                string rmrk = tcRow.ItemArray[5].ToString();
                                string attch = tcRow.ItemArray[6].ToString();
                                string htmlattch = tcRow.ItemArray[7].ToString();
                                string com = tcRow.ItemArray[9].ToString();
                                string attachurl = tcRow.ItemArray[10].ToString();

                                if (string.IsNullOrWhiteSpace(attch) == false)
                                {
                                    try
                                    {
                                        attch = attch.Substring(0, attch.IndexOf('\\')) + rootFolderExt +
                                                attch.Substring(attch.IndexOf('\\'), attch.Length - attch.IndexOf('\\'));
                                    }
                                    catch
                                    {
                                        // ignored
                                    }
                                }
                                if (string.IsNullOrWhiteSpace(htmlattch) == false)
                                {
                                    try
                                    {
                                        htmlattch = htmlattch.Substring(0, htmlattch.IndexOf('\\')) + rootFolderExt +
                                                    htmlattch.Substring(htmlattch.IndexOf('\\'),
                                                                        htmlattch.Length - htmlattch.IndexOf('\\'));
                                    }
                                    catch
                                    {
                                        // ignored
                                    }
                                }
                                int outParam;
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
                                catch
                                {
                                    // ignored
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

                                finalHtml.Append(populatedTr);

                                // updating each TestStep Status Tag Style
                                if (stats == ExecutionStatus.Fail)
                                {
                                    Property.TotalStepFail++;

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
                                    Property.TotalStepWarning++;

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
                                    Property.TotalStepPass++;
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
                            
                            finalHtml.Append(newTestCaseBodyEnd);

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

                //using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
                //{
                //    sw.WriteLine("Logfile outside Foreach ");
                //}
                finalHtml = finalHtml.Replace("$TotalCaseExecuted", Property.TotalCaseExecuted.ToString());
                finalHtml = finalHtml.Replace("$TotalCasePass", Property.TotalCasePass.ToString());
                finalHtml = finalHtml.Replace("$TotalCaseFail", Property.TotalCaseFail.ToString());
                finalHtml = finalHtml.Replace("$TotalCaseWarning", Property.TotalCaseWarning.ToString());
                // Ending Html Report
                finalHtml.Append(htmlEnd);
                // Adding Complete Suit execution time

                finalHtml = finalHtml.Replace("$ExecutionStartTime", executionStartTime.ToString(Utility.GetParameter("DateTimeFormat")));
                finalHtml = finalHtml.Replace("$ExecutionEndTime", executionEndTime.ToString(Utility.GetParameter("DateTimeFormat")));

                TimeSpan time = executionEndTime - executionStartTime;

                Property.ExecutionStartDateTime = executionStartTime.ToString(Utility.GetParameter("DateTimeFormat"));
                Property.ExecutionEndDateTime = executionEndTime.ToString(Utility.GetParameter("DateTimeFormat"));
                Utility.SetVariable("ExecutionStartDateTime", Property.ExecutionStartDateTime);
                Utility.SetVariable("ExecutionEndDateTime", Property.ExecutionEndDateTime);


                finalHtml = finalHtml.Replace("$totalRunDuration", time.ToString());

                Property.ReportSummaryBody = Convert.ToString(finalHtml);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.StackTrace.Substring(exception.StackTrace.LastIndexOf(' ')));
                Console.WriteLine(exception.Message);
            }
        }
    }

}
