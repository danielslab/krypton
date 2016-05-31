/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: BrowserManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Create HTML Report Template and Set Up the Parameters
*****************************************************************************/
using System;
using System.Text;
using System.IO;
using Common;
using System.Data;
using System.Reflection;
using Ionic.Zip;
using System.Drawing;
using System.Drawing.Imaging;
using PreMailer.Net;

namespace Reporting
{
    public class HtmlReport
    {
        static readonly string ThinkSysLogo = ConsoleMessages.THINKSYSLOGO;
        public static void CreateHtmlReport(string htmlFilename, bool zipRequired, bool logoUpload, bool sauceFlag, bool isSummaryRequiredInResults, bool isGrid, bool callFromKryptonVbScriptGrid, bool creatHtmlReport)
        {
            Property.TotalCaseExecuted = 0;
            Property.TotalCasePass = 0;
            Property.TotalCaseFail = 0;
            Property.TotalCaseWarning = 0;

            Property.TotalStepExecuted = 0;
            Property.TotalStepFail = 0;
            Property.TotalStepPass = 0;
            Property.TotalStepWarning = 0;

            FileInfo oFileInfo = new FileInfo(Assembly.GetExecutingAssembly().Location);
            if (oFileInfo.Directory != null)
            {
                string templatePath = Path.Combine(oFileInfo.Directory.FullName, "ReportTemplate");

                string reportTemplate = Path.Combine(templatePath, "HtmlReportTemplate.htm");
                StreamReader sr = new StreamReader(reportTemplate);
                string htmlReportTemplate = sr.ReadToEnd();
                sr.Close();

                string parentRow = Path.Combine(templatePath, "ParentRow.htm");
                sr = new StreamReader(parentRow);
                string parentRowTemplate = sr.ReadToEnd();
                sr.Close();

                if (string.IsNullOrWhiteSpace(htmlFilename))
                {
                    htmlFilename = "HtmlReport.html";
                }
                Property.Time = "HH:mm:ss";

                DateTime executionStartTime = DateTime.Now;
                DateTime executionEndTime = DateTime.Now;
                DateTime currentTime = executionEndTime;

                string[] xmlFilesLocation = LogFile.AllXmlFilesLocation.Split(';');
                string htmlFileLocation = Property.HtmlFileLocation + "/" + htmlFilename;

                Property.FinalExecutionStatus = ExecutionStatus.Pass;
                int count = 0;
                StringBuilder sbHtml = new StringBuilder();
                StringBuilder sbSummaryHtml = new StringBuilder();
                StringBuilder sbInlineHtml = new StringBuilder();
                //Somtimes if multiple testcases are running then browser name comes blank after first test case. So set same browser name for next textcases too.
                string lastBrowserName = string.Empty;
                foreach (string xmlLocation in xmlFilesLocation)
                {
                    string tcName = string.Empty;
                    string tcId = string.Empty;
                    string tcBrowser = string.Empty;
                    string tcMachine = string.Empty;
                    string tcStartTime = string.Empty;
                    string tcEndTime = string.Empty;
                    string tcDuration = string.Empty;

                    if (File.Exists(xmlLocation))
                    {
                        count++;
                        var rootFolderExt = Property.DateTime;

                        Property.TotalCaseExecuted++;
                        DataTable dtTemp = null;
                        try
                        {
                            DataSet dsStore = new DataSet();
                            // copying all data from XML to datatable
                            dsStore.ReadXml(xmlLocation);
                            dtTemp = dsStore.Tables[0];
                            // getting TestCase Name and ID from extended properties
                            if (dtTemp.ExtendedProperties["TestCase Name"] != null)
                                tcName = dtTemp.ExtendedProperties["TestCase Name"].ToString();
                            if (dtTemp.ExtendedProperties["TestCase Id"] != null)
                                tcId = dtTemp.ExtendedProperties["TestCase Id"].ToString();
                            if (dtTemp.ExtendedProperties["Browser"] != null)
                                tcBrowser = dtTemp.ExtendedProperties["Browser"].ToString();
                            if (dtTemp.ExtendedProperties["RCMachineId"] != null)
                                tcMachine = dtTemp.ExtendedProperties["RCMachineId"].ToString();
                            //Extract date/time information from xml files
                            if (dtTemp.ExtendedProperties["ExecutionEndTime"] != null)
                                tcEndTime = dtTemp.ExtendedProperties["ExecutionEndTime"].ToString();
                            if (dtTemp.ExtendedProperties["ExecutionStartTime"] != null)
                                tcStartTime = dtTemp.ExtendedProperties["ExecutionStartTime"].ToString();
                            if (dtTemp.ExtendedProperties["ExecutionDuration"] != null)
                                tcDuration = dtTemp.ExtendedProperties["ExecutionDuration"].ToString();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

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
                        string tempParentRow = parentRowTemplate;
                        tempParentRow = tempParentRow.Replace("$RowId$", "RowId" + count);
                        tempParentRow = tempParentRow.Replace("$Row-Childs-Id$", "RowChildsId" + count);
                        tempParentRow = tempParentRow.Replace("$TestCaseId$", tcId);
                        if (tcName.Trim().Length > 0)
                            tempParentRow = tempParentRow.Replace("$TestCaseName$", " - " + tcName);
                        else
                            tempParentRow = tempParentRow.Replace("$TestCaseName$", string.Empty);

                        if (tcBrowser.Trim().Length > 0)
                            lastBrowserName = tcBrowser;
                        tempParentRow = tempParentRow.Replace("$Browser$", lastBrowserName);
                        tempParentRow = tempParentRow.Replace("$Machine$", tcMachine);
                        tempParentRow = tcDuration.Trim().Length == 0 ? tempParentRow.Replace("$TotalTestcaseDuration$", dtTemp.Rows[0][3] + " : " + dtTemp.Rows[0][4]) : tempParentRow.Replace("$TotalTestcaseDuration$", tcDuration);

                        string headerRemarkFail = string.Empty;
                        string headerRemarkWarning = string.Empty;
                        string finalStatus = ExecutionStatus.Pass;
                        string childRows = GetChildRowsHtml(dtTemp, rootFolderExt, ref finalStatus, ref headerRemarkFail, ref headerRemarkWarning);

                        tempParentRow = tempParentRow.Replace("$FinalStatus$", finalStatus);
                        if (finalStatus == ExecutionStatus.Fail)
                        {
                            tempParentRow = tempParentRow.Replace("$Remarks$", headerRemarkFail);
                            tempParentRow = tempParentRow.Replace("$FinalStatusClass$", "Fail");
                            Property.TotalCaseFail++;
                            Property.ExecutionFailReason = headerRemarkFail;
                        }
                        else if (finalStatus == ExecutionStatus.Pass)
                        {
                            tempParentRow = tempParentRow.Replace("$Remarks$", string.Empty);
                            tempParentRow = tempParentRow.Replace("$FinalStatusClass$", "Pass");
                            Property.TotalCasePass++;

                        }
                        else
                        {
                            tempParentRow = tempParentRow.Replace("$Remarks$", headerRemarkWarning);
                            tempParentRow = tempParentRow.Replace("$FinalStatusClass$", "Warning");
                            Property.TotalCaseWarning++;
                        }

                        string parentRowComplete = tempParentRow.Replace("$HTMLChildRows$", childRows);
                        sbHtml.Append(parentRowComplete);

                        string inlineHtml = parentRowComplete;
                        sbInlineHtml.Append(RemoveExpandIcon(inlineHtml, "[" + count + "]", true));

                        //Remove child which include step actions.
                        string tempParentSummaryRow = tempParentRow;
                        string rowDeliminetor = "</tr>";
                        int index = tempParentSummaryRow.IndexOf(rowDeliminetor, StringComparison.Ordinal);
                        tempParentSummaryRow = tempParentSummaryRow.Substring(0, index + rowDeliminetor.Length);
                        sbSummaryHtml.Append(RemoveExpandIcon(tempParentSummaryRow, "[" + count + "]", false));
                    }
                    using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
                    {
                        sw.WriteLine(count);
                    }
                }
                using (StreamWriter sw = new StreamWriter(Property.ErrorLog))
                {
                    sw.WriteLine("Foreach end of HTMlReport..");
                }
                htmlReportTemplate = htmlReportTemplate.Replace("$KryptonVersion$", Property.KryptonVersion);
                htmlReportTemplate = htmlReportTemplate.Replace("$projectname$", Utility.GetParameter("ProjectName").Trim());
                htmlReportTemplate = htmlReportTemplate.Replace("$TestSuitName$", Utility.GetParameter("TestSuite").Trim());
                htmlReportTemplate = htmlReportTemplate.Replace("$Environment$", Utility.GetParameter(Property.Environment));
                htmlReportTemplate = htmlReportTemplate.Replace("$ExecutionStartedAt$", Property.ExecutionStartDateTime);
                htmlReportTemplate = htmlReportTemplate.Replace("$ExecutionFinishedAt$", Property.ExecutionEndDateTime);
                TimeSpan time = executionEndTime - executionStartTime;
                htmlReportTemplate = htmlReportTemplate.Replace("$TotalDuration$", time.ToString());

                htmlReportTemplate = htmlReportTemplate.Replace("$TestCase(s)Executed$", Property.TotalCaseExecuted.ToString());
                htmlReportTemplate = htmlReportTemplate.Replace("$TotalPassed$", Property.TotalCasePass.ToString());
                htmlReportTemplate = htmlReportTemplate.Replace("$TotalFailed$", Property.TotalCaseFail.ToString());
                htmlReportTemplate = htmlReportTemplate.Replace("$TotalWarning$", Property.TotalCaseWarning.ToString());
                htmlReportTemplate = htmlReportTemplate.Replace("$CompanyLogo$", ImageToBase64(Property.CompanyLogo));
                string finalHtmlReportTemplate = htmlReportTemplate.Replace("$TableBody$", sbHtml.ToString());
                if (creatHtmlReport)
                    WriteInFile(htmlFileLocation, finalHtmlReportTemplate);

                string inlineHtmlTemplate = RemoveImagesAndHeader(htmlReportTemplate.Replace("$TableBody$", sbInlineHtml.ToString()));
                InlineResult inlineCss1 = PreMailer.Net.PreMailer.MoveCssInline(inlineHtmlTemplate);
                if (isGrid)
                    WriteInFile(htmlFileLocation.Replace(".html", "smail.html"), inlineCss1.Html);
                else
                    WriteInFile(htmlFileLocation.Replace(".html", "mail.html"), inlineCss1.Html); // only used for sending mail through krypton

                string htmlSummaryReportTemplate = htmlReportTemplate;
                htmlSummaryReportTemplate = htmlSummaryReportTemplate.Replace("$TableBody$", sbSummaryHtml.ToString());
                htmlFileLocation = htmlFileLocation.Replace(".html", "s.html");

                #region
                if (isSummaryRequiredInResults && creatHtmlReport)
                    WriteInFile(htmlFileLocation, htmlSummaryReportTemplate);
                #endregion
                // created in result folder with logo and report header

                #region
                htmlSummaryReportTemplate = RemoveImagesAndHeader(htmlSummaryReportTemplate);
                InlineResult inlineCss3 = PreMailer.Net.PreMailer.MoveCssInline(htmlSummaryReportTemplate);
                Property.ReportSummaryBody = inlineCss3.Html;
                if (callFromKryptonVbScriptGrid)
                    WriteInFile(htmlFileLocation, inlineCss3.Html);
                #endregion
                // for Mail Summary and Existing Grid body without logo and report header

                if (zipRequired) //if ziprequied flag is ture then create zip file with all images and html file
                {
                    CreateZip(xmlFilesLocation);
                }
            }
        }

        static string RemoveExpandIcon(string htmlString, string replacedString, bool isAppend)
        {
            int fIndex = htmlString.IndexOf("<div", StringComparison.Ordinal);
            int lIndex = htmlString.IndexOf("</div", StringComparison.Ordinal);
            string tempDivRow = htmlString.Substring(fIndex, lIndex - fIndex + 6);
            if (isAppend)
            {
                replacedString = tempDivRow.Replace("</div", replacedString + "</div");
            }
            htmlString = htmlString.Replace(tempDivRow, replacedString);
            return htmlString;
        }

        static string RemoveImagesAndHeader(string htmlSummaryReportTemplate)
        {
            int fIndex1 = htmlSummaryReportTemplate.IndexOf("<div class=\"ReportTopTitle\">", StringComparison.Ordinal);
            int lIndex1 = htmlSummaryReportTemplate.IndexOf("</div>", fIndex1, StringComparison.Ordinal);
            string tempDivRow1 = htmlSummaryReportTemplate.Substring(fIndex1, lIndex1 - fIndex1 + 6);
            htmlSummaryReportTemplate = htmlSummaryReportTemplate.Replace(tempDivRow1, string.Empty);

            fIndex1 = htmlSummaryReportTemplate.IndexOf("<img", StringComparison.Ordinal);
            lIndex1 = htmlSummaryReportTemplate.IndexOf("/>", fIndex1, StringComparison.Ordinal);
            tempDivRow1 = htmlSummaryReportTemplate.Substring(fIndex1, lIndex1 - fIndex1 + 3);
            htmlSummaryReportTemplate = htmlSummaryReportTemplate.Replace(tempDivRow1, string.Empty);

            htmlSummaryReportTemplate = htmlSummaryReportTemplate.Replace("Report Details", "<hr width=\"100%;\" size=\"2\"> Krypton Automation Report");
            return htmlSummaryReportTemplate;
        }

        static void WriteInFile(string htmlFileLocation, string htmlRerpotText)
        {
            if (File.Exists(htmlFileLocation))
                File.Delete(htmlFileLocation);
            StreamWriter sWriter = new StreamWriter(htmlFileLocation, false, Encoding.UTF8);
            // Saving Final Html Report
            sWriter.WriteLine(htmlRerpotText);
            sWriter.Close();
        }

        static string ImageToBase64(string logoImgPath)
        {
            string base64String = ThinkSysLogo;
            try
            {
                if (File.Exists(logoImgPath))
                {
                    ImageFormat format = null;
                    var extension = Path.GetExtension(logoImgPath);
                    if (extension != null)
                        switch (extension.ToLower())
                        {
                            case ".bmp":
                                format = ImageFormat.Bmp;
                                base64String = "data:image/bmp;base64,";
                                break;
                            case ".gif":
                                format = ImageFormat.Gif;
                                base64String = "data:image/gif;base64,";
                                break;
                            case ".jpeg":
                            case ".jpg":
                                format = ImageFormat.Jpeg;
                                base64String = "data:image/jpeg;base64,";
                                break;
                            default:
                                format = ImageFormat.Png;
                                base64String = "data:image/png;base64,";
                                break;
                        }
                    Image logoImage;
                    using (FileStream stream = new FileStream(logoImgPath, FileMode.Open, FileAccess.Read))
                    {
                        logoImage = Image.FromStream(stream);
                    }
                    using (MemoryStream ms = new MemoryStream())
                    {
                        // Convert Image to byte[]
                        if (format != null) logoImage.Save(ms, format);
                        byte[] imageBytes = ms.ToArray();
                        // Convert byte[] to Base64 String
                        base64String += Convert.ToBase64String(imageBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                base64String = ThinkSysLogo;
                Console.WriteLine(ConsoleMessages.COULD_NOT_READ_PROJECT_LOGO);
                Console.WriteLine("Error: " + ex.Message);
            }
            return base64String;
        }

        static string UpdateAttachment(string attch, string rootFolderExt)
        {
            if (string.IsNullOrWhiteSpace(attch) == false && attch.IndexOf('\\') > -1)
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
            return attch;
        }
        static string GetChildRowsHtml(DataTable dtTemp, string rootFolderExt, ref string finalStatus, ref string headerRemarkFail, ref string headerRemarkWarning)
        {
            StringBuilder sbChildHtmlRows = new StringBuilder();
            FileInfo oFileInfo = new FileInfo(Assembly.GetExecutingAssembly().Location);
            if (oFileInfo.Directory != null)
            {
                string templatePath = Path.Combine(oFileInfo.Directory.FullName, "ReportTemplate");
                string childRow = Path.Combine(templatePath, "ChildRow.htm");
                StreamReader sr = new StreamReader(childRow);
                string childRowTemplate = sr.ReadToEnd();
                sr.Close();

                int rowNo = 0;
                bool firstFail = true;
                bool firstWarning = true;
                foreach (DataRow tcRow in dtTemp.Rows) // Loop over the rows.
                {
                    string tempChildRowTemplate = childRowTemplate;
                    Property.TotalStepExecuted++;
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$AltStyle$", rowNo % 2 == 0 ? string.Empty : "class=\"alt\"");

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

                    attch = UpdateAttachment(attch, rootFolderExt);
                    htmlattch = UpdateAttachment(htmlattch, rootFolderExt);

                    tempChildRowTemplate = tempChildRowTemplate.Replace("$StepNo$", stpnmb);
                    // update comments
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$Child-Steps$", string.IsNullOrWhiteSpace(com) == false ? com : stpdesp);

                    string tempStatusWithAttach = string.Empty;
                    // if Attachment (Screenshot) is present, Status should act as link
                    if (string.IsNullOrWhiteSpace(attch) == false)
                        tempStatusWithAttach = "<a target='_blank' title='" + attachurl + "' href='" + attch + "'>" + stats + "</a>&nbsp;";
                    if (string.IsNullOrWhiteSpace(htmlattch) == false)
                        tempStatusWithAttach += "<a target='_blank' href='" + htmlattch + "'>(html)</a>";
                    if (tempStatusWithAttach.Trim().Length == 0)
                        tempStatusWithAttach = stats;

                    // Update Status
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$Child-Status$", tempStatusWithAttach);

                    // Update Execution Date and Time
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$Child-ExecutionTime$", exedt + "" + exetm);

                    // Update Remarks
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$Child-Remarks$", rmrk);

                    // updating each TestStep Status Tag Style
                    if (stats == ExecutionStatus.Fail)
                    {
                        Property.TotalStepFail++;
                        tempChildRowTemplate = tempChildRowTemplate.Replace("$StepStatus$", "Fail");

                        // If Current TestCase Step is fail (means complete TestCase status will be fail)
                        finalStatus = ExecutionStatus.Fail;
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
                        tempChildRowTemplate = tempChildRowTemplate.Replace("$StepStatus$", "Warning");
                        if (finalStatus != ExecutionStatus.Fail)
                            finalStatus = ExecutionStatus.Warning;
                        //Added remarks in test case header if test case has encountered a warning step.
                        if (firstWarning)
                        {
                            headerRemarkWarning = rmrk;
                            firstWarning = false;
                        }
                    }
                    else
                    {
                        tempChildRowTemplate = tempChildRowTemplate.Replace("$StepStatus$", "Pass");
                        Property.TotalStepPass++;
                    }
                    sbChildHtmlRows.Append(tempChildRowTemplate);
                    rowNo++;
                }
            }
            return sbChildHtmlRows.ToString();
        }
        static void CreateZip(string[] xmlFilesLocation)
        {
            using (ZipFile zip = new ZipFile())
            {
                foreach (string xmlLocation in xmlFilesLocation)
                {
                    // Path to directory of files to compress and decompress.                               
                    string directory = Path.GetDirectoryName(xmlLocation);
                    if (directory != null && Directory.Exists(directory))
                    {
                        FileInfo fileInfo = new FileInfo(xmlLocation);
                        if (fileInfo.Directory != null)
                            if (fileInfo.Directory.Parent != null)
                                zip.AddDirectory(directory, fileInfo.Directory.Parent.Name + "/" + fileInfo.Directory.Name);
                    }
                }
                zip.AddFile(Property.HtmlFileLocation + "\\HtmlReport.html", "./");
                if (File.Exists(Property.HtmlFileLocation + "\\" + Property.ReportZipFileName))
                    //delete zip file if already exists
                    File.Delete(Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);
                zip.Save(Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);
            }
        }
    }
}
