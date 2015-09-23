/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: BrowserManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Create HTML Report Template and Set Up the Parameters
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
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
    public class HTMLReport
    {
        const string _ThinkSysLogo = @"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALwAAAA1CAMAAADMKHGJAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAPBQTFRFAAAAAK7vCT1g////CT1gCT1gCT1gAK7vCT1gAK7vCT1gAK7vCT1gAK7vCT1gAK7vAK7vCT1gAK7vCT1gCT1gAK7vCT1gCT1gAK7vCT1gCT1gAK7vCT1gAK7vAK7vCT1gAK7vAK7vCT1gCT1gCT1gAK7vCT1gAK7vCT1gCT1gAK7vCT1gCT1gAK7vCT1gAK7vCT1gCT1gAK7vCT1gAK7vCT1gCT1gAK7vAK7vCT1gCT1gAK7vCT1gCT1gCT1gAK7vAK7vCT1gCT1gAK7vCT1gCT1gAK7vCT1gCT1gAK7vCT1gCT1gAK7vCT1gAK7vCT1gYRXGaAAAAE50Uk5TAAAAAAMJDBAQEhIeHiAgISotMDA8QEBLUFBUWlpdYGBpcHB1eICAgYeKj4+WmZmfn6Worq+vtLe/v8DDw8bJzM/P0t/f4eTk7e/v9vz8KLY2iAAAAwNJREFUaN7t2H170jAQAPBYnMUp6FphTCdMttbJVFBBmc4XLL7AXPn+38bkjiZpmrasbNA+T++vW9orv+y5pm1IpcBBSnyJL/EFwtvNaNQJIXXIiAh+pApZFUdtTalarV6q1ohEjQ6bkJniPAsGln9EaxoVcr6IhktPdiGT8PxIEzLAuPNI7dQmavWE5aPgOi3Pj0ab8SBriF8cwwDLzL6mxlsP39HULuZVpdqGvBP8231dmMl4nd3vrYcf6fBwRK4+41Ni8ULnGJJk/ExXZK2HP9fiu0r1lKWD4DKOH9M1SXhdjUduA69UY9ccJOLNDPgexXdciAH8xAD/aF4HPxFLTVeHx64hYXx42bDIKvh+uMik+GAhFEsIxsr4c2k11eHnoa5Z4okm0vCOWnD7+APlwoXCQz9OSTHx0DVnxcRj19hbweuWwdXxIkb1pKXSa2VYKvvmpvCLuZ20zvvW9fH+cGP4xSQR72TA+5vDB9e2UvB9h4fH8cNt44PnVC8ZHwmYsZcF74rIgA9K8R0oeK+stR05VsHTg6Ga8Ur4G3g94ONdog2O1zfUTFfjbBCPb8XTFDz5ocOPt47vqFfX4mvjXOKryqulHq/GOB94fD3jt2zB8HV+y8Y/pDA8Xdtke0jdFB5PnabjLT+HeP5NkoLv5RHP9xD0+HZc1/j9WLy3Oby7vGWd2H0bbdd4Zhx+1uD4OjzH6xIEBuRfD/YVxKmw89CRbkt+TqSaVGHEVp7zGJb4IFGOtHDrUlPTZrsH5RZ3ic+Cv5OfqFQMEXIeEwz/+Hsu4iXg7/6FZekn4D/gd8xDmh5iemxIOcMf+bmIZ4B/grBXhpjILwZ+h+N7cs7w73Nh/3cf8KcI2zfERF4z8B8xD54z/O9c4D9hz38F2OUOw58g/in17i0/hA05p/hH+eiaI8DvXgHsM9ywF5Bf7VLwMYIPDTmn+Of5wD8A/PJePGH4eziRCwb+KOYhcop/kwv7F1wq3yLeZvh9zE+pd+cS0m9GKC+fsCW+xJf4Ep8a/wFiva9qhW3zQwAAAABJRU5ErkJggg==";
        public static void CreateHtmlReport(string htmlFilename, bool zipRequired, bool logoUpload, bool sauceFlag, bool IsSummaryRequiredInResults, bool IsGrid, bool CallFromKryptonVBScriptGrid, bool creatHtmlReport)
        {
            Common.Property.TotalCaseExecuted = 0;
            Common.Property.TotalCasePass = 0;
            Common.Property.TotalCaseFail = 0;
            Common.Property.TotalCaseWarning = 0;

            Common.Property.TotalStepExecuted = 0;
            Common.Property.TotalStepFail = 0;
            Common.Property.TotalStepPass = 0;
            Common.Property.TotalStepWarning = 0;

            FileInfo oFileInfo = new FileInfo(Assembly.GetExecutingAssembly().Location);
            string TemplatePath = Path.Combine(oFileInfo.Directory.FullName, "ReportTemplate");

            string reportTemplate = Path.Combine(TemplatePath, "HtmlReportTemplate.htm");
            StreamReader sr = new StreamReader(reportTemplate);
            string HtmlReportTemplate = sr.ReadToEnd();
            sr.Close();

            string parentRow = Path.Combine(TemplatePath, "ParentRow.htm");
            sr = new StreamReader(parentRow);
            string ParentRowTemplate = sr.ReadToEnd();
            sr.Close();

            string childRow = Path.Combine(TemplatePath, "ChildRow.htm");
            sr = new StreamReader(childRow);
            string ChildRowTemplate = sr.ReadToEnd();
            sr.Close();

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

            string[] xmlFilesLocation = LogFile.allXmlFilesLocation.Split(';');
            string htmlFileLocation = Common.Property.HtmlFileLocation + "/" + htmlFilename;

            Property.FinalExecutionStatus = ExecutionStatus.Pass;
            int Count = 0;
            StringBuilder sbHtml = new StringBuilder();
            StringBuilder sbSummaryHtml = new StringBuilder();
            StringBuilder sbInlineHtml = new StringBuilder();
            //Somtimes if multiple testcases are running then browser name comes blank after first test case. So set same browser name for next textcases too.
            string lastBrowserName = string.Empty;
            foreach (string xmlLocation in xmlFilesLocation)
            {

                string tcName = string.Empty;
                string tcID = string.Empty;
                string tcBrowser = string.Empty;
                string tcMachine = string.Empty;
                string tcStartTime = string.Empty;
                string tcEndTime = string.Empty;
                string tcDuration = string.Empty;
                string tcJobResultLocation = @"javascript:void(0)";

                if (File.Exists(xmlLocation))
                {
                    Count++;
                    string rootFolderName = new FileInfo(xmlLocation).Directory.Parent.Name;
                    string rootFolderExt = string.Empty;
                    if (rootFolderName.IndexOf('-') > -1)
                        rootFolderExt = rootFolderName.Substring(rootFolderName.IndexOf('-'), rootFolderName.Length - rootFolderName.IndexOf('-'));


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
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }

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
                    string tempParentRow = ParentRowTemplate;
                    tempParentRow = tempParentRow.Replace("$RowId$", "RowId" + Count);
                    tempParentRow = tempParentRow.Replace("$Row-Childs-Id$", "RowChildsId" + Count);
                    tempParentRow = tempParentRow.Replace("$TestCaseId$", tcID);
                    if (tcName.Trim().Length > 0)
                        tempParentRow = tempParentRow.Replace("$TestCaseName$", " - " + tcName);
                    else
                        tempParentRow = tempParentRow.Replace("$TestCaseName$", string.Empty);

                    if (tcBrowser.Trim().Length > 0)
                        lastBrowserName = tcBrowser;
                    tempParentRow = tempParentRow.Replace("$Browser$", lastBrowserName);
                    tempParentRow = tempParentRow.Replace("$Machine$", tcMachine);
                    if (tcDuration.Trim().Length == 0)
                    {
                        tempParentRow = tempParentRow.Replace("$TotalTestcaseDuration$", dtTemp.Rows[0][3].ToString() + " : " + dtTemp.Rows[0][4].ToString());
                    }
                    else
                    {
                        tempParentRow = tempParentRow.Replace("$TotalTestcaseDuration$", tcDuration);
                    }

                    string headerRemarkFail = string.Empty;
                    string headerRemarkWarning = string.Empty;
                    string finalStatus = ExecutionStatus.Pass;
                    string ChildRows = GetChildRowsHtml(dtTemp, rootFolderExt, ref finalStatus, ref headerRemarkFail, ref headerRemarkWarning);

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

                    string parentRowComplete = tempParentRow.Replace("$HTMLChildRows$", ChildRows);
                    sbHtml.Append(parentRowComplete);

                    string inlineHtml = parentRowComplete;
                    sbInlineHtml.Append(RemoveExpandIcon(inlineHtml, "[" + Count + "]", true));

                    //Remove child which include step actions.
                    string tempParentSummaryRow = tempParentRow;
                    string rowDeliminetor = "</tr>";
                    int index = tempParentSummaryRow.IndexOf(rowDeliminetor);
                    tempParentSummaryRow = tempParentSummaryRow.Substring(0, index + rowDeliminetor.Length);
                    sbSummaryHtml.Append(RemoveExpandIcon(tempParentSummaryRow, "[" + Count + "]", false));


                }
            }
            HtmlReportTemplate = HtmlReportTemplate.Replace("$KryptonVersion$", Property.KRYPTONVERSION);
            HtmlReportTemplate = HtmlReportTemplate.Replace("$projectname$", Common.Utility.GetParameter("ProjectName").Trim());
            HtmlReportTemplate = HtmlReportTemplate.Replace("$TestSuitName$", Common.Utility.GetParameter("TestSuite").Trim());
            HtmlReportTemplate = HtmlReportTemplate.Replace("$Environment$", Common.Utility.GetParameter(Common.Property.ENVIRONMENT));
            HtmlReportTemplate = HtmlReportTemplate.Replace("$ExecutionStartedAt$", Common.Property.ExecutionStartDateTime);
            HtmlReportTemplate = HtmlReportTemplate.Replace("$ExecutionFinishedAt$", Common.Property.ExecutionEndDateTime);
            TimeSpan time = executionEndTime - executionStartTime;
            HtmlReportTemplate = HtmlReportTemplate.Replace("$TotalDuration$", time.ToString());

            HtmlReportTemplate = HtmlReportTemplate.Replace("$TestCase(s)Executed$", Common.Property.TotalCaseExecuted.ToString());
            HtmlReportTemplate = HtmlReportTemplate.Replace("$TotalPassed$", Common.Property.TotalCasePass.ToString());
            HtmlReportTemplate = HtmlReportTemplate.Replace("$TotalFailed$", Common.Property.TotalCaseFail.ToString());
            HtmlReportTemplate = HtmlReportTemplate.Replace("$TotalWarning$", Common.Property.TotalCaseWarning.ToString());

            HtmlReportTemplate = HtmlReportTemplate.Replace("$CompanyLogo$", ImageToBase64(Common.Property.CompanyLogo));

            string finalHtmlReportTemplate = HtmlReportTemplate.Replace("$TableBody$", sbHtml.ToString());
            if (creatHtmlReport)
            WriteInFile(htmlFileLocation, finalHtmlReportTemplate);

            string inlineHtmlTemplate = RemoveImagesAndHeader(HtmlReportTemplate.Replace("$TableBody$", sbInlineHtml.ToString()));
            InlineResult inlineCss1 = PreMailer.Net.PreMailer.MoveCssInline(inlineHtmlTemplate, false);


            if (IsGrid)
                WriteInFile(htmlFileLocation.Replace(".html", "smail.html"), inlineCss1.Html);
            else
                WriteInFile(htmlFileLocation.Replace(".html", "mail.html"), inlineCss1.Html); // only used for sending mail through krypton

            string HtmlSummaryReportTemplate = HtmlReportTemplate;
            HtmlSummaryReportTemplate = HtmlSummaryReportTemplate.Replace("$TableBody$", sbSummaryHtml.ToString());
            htmlFileLocation = htmlFileLocation.Replace(".html", "s.html");

            #region
            if (IsSummaryRequiredInResults && creatHtmlReport)
                WriteInFile(htmlFileLocation, HtmlSummaryReportTemplate);
            #endregion
            // created in result folder with logo and report header

            #region
            HtmlSummaryReportTemplate = RemoveImagesAndHeader(HtmlSummaryReportTemplate);
            InlineResult inlineCss3 = PreMailer.Net.PreMailer.MoveCssInline(HtmlSummaryReportTemplate, false);
            Common.Property.ReportSummaryBody = inlineCss3.Html;
            if (CallFromKryptonVBScriptGrid)
                WriteInFile(htmlFileLocation, inlineCss3.Html);
            #endregion
            // for Mail Summary and Existing Grid body without logo and report header

            if (zipRequired) //if ziprequied flag is ture then create zip file with all images and html file
            {
                CreateZip(xmlFilesLocation);
            }
        }

        static string RemoveExpandIcon(string HtmlString, string ReplacedString, bool IsAppend)
        {
            int fIndex = HtmlString.IndexOf("<div");
            int lIndex = HtmlString.IndexOf("</div");
            string tempDivRow = HtmlString.Substring(fIndex, lIndex - fIndex + 6);
            if (IsAppend)
            {
                ReplacedString = tempDivRow.Replace("</div", ReplacedString + "</div");
            }
            HtmlString = HtmlString.Replace(tempDivRow, ReplacedString);
            return HtmlString;
        }

        static string RemoveImagesAndHeader(string HtmlSummaryReportTemplate)
        {

            int fIndex1 = HtmlSummaryReportTemplate.IndexOf("<div class=\"ReportTopTitle\">");
            int lIndex1 = HtmlSummaryReportTemplate.IndexOf("</div>", fIndex1);
            string tempDivRow1 = HtmlSummaryReportTemplate.Substring(fIndex1, lIndex1 - fIndex1 + 6);
            HtmlSummaryReportTemplate = HtmlSummaryReportTemplate.Replace(tempDivRow1, string.Empty);

            fIndex1 = HtmlSummaryReportTemplate.IndexOf("<img");
            lIndex1 = HtmlSummaryReportTemplate.IndexOf("/>", fIndex1);
            tempDivRow1 = HtmlSummaryReportTemplate.Substring(fIndex1, lIndex1 - fIndex1 + 3);
            HtmlSummaryReportTemplate = HtmlSummaryReportTemplate.Replace(tempDivRow1, string.Empty);

            HtmlSummaryReportTemplate = HtmlSummaryReportTemplate.Replace("Report Details", "<hr width=\"100%;\" size=\"2\"> Krypton Automation Report");
            return HtmlSummaryReportTemplate;
        }

        static void WriteInFile(string htmlFileLocation, string HtmlRerpotText)
        {
            if (File.Exists(htmlFileLocation))
                File.Delete(htmlFileLocation);
            StreamWriter sWriter = new StreamWriter(htmlFileLocation, false, Encoding.UTF8);
            // Saving Final Html Report
            sWriter.WriteLine(HtmlRerpotText);
            sWriter.Close();
        }

        static string ImageToBase64(string LogoImgPath)
        {
            string base64String = _ThinkSysLogo;
            try
            {
                if (File.Exists(LogoImgPath))
                {
                    ImageFormat format = ImageFormat.Png;
                    switch (Path.GetExtension(LogoImgPath).ToLower())
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
                    Image LogoImage = null;
                    using (FileStream stream = new FileStream(LogoImgPath, FileMode.Open, FileAccess.Read))
                    {
                        LogoImage = Image.FromStream(stream);
                    }
                    using (MemoryStream ms = new MemoryStream())
                    {
                        // Convert Image to byte[]
                        LogoImage.Save(ms, format);
                        byte[] imageBytes = ms.ToArray();
                        // Convert byte[] to Base64 String
                        base64String += Convert.ToBase64String(imageBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                base64String = _ThinkSysLogo;
                Console.WriteLine("Warning: Coluld not read project logo file. Using default logo of Krypton Report.");
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
                catch { }
            }
            return attch;
        }
        static string GetChildRowsHtml(DataTable dtTemp, string rootFolderExt, ref string finalStatus, ref string headerRemarkFail, ref string headerRemarkWarning)
        {
            StringBuilder sbChildHtmlRows = new StringBuilder();
            FileInfo oFileInfo = new FileInfo(Assembly.GetExecutingAssembly().Location);
            string TemplatePath = Path.Combine(oFileInfo.Directory.FullName, "ReportTemplate");
            string childRow = Path.Combine(TemplatePath, "ChildRow.htm");
            StreamReader sr = new StreamReader(childRow);
            string ChildRowTemplate = sr.ReadToEnd();
            sr.Close();

            int rowNo = 0;
            bool firstFail = true;
            bool firstWarning = true;
            foreach (DataRow tcRow in dtTemp.Rows) // Loop over the rows.
            {
                string tempChildRowTemplate = ChildRowTemplate;
                Common.Property.TotalStepExecuted++;
                if (rowNo % 2 == 0)
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$AltStyle$", string.Empty);
                else
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$AltStyle$", "class=\"alt\"");

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

                attch = UpdateAttachment(attch, rootFolderExt);
                htmlattch = UpdateAttachment(htmlattch, rootFolderExt);

                tempChildRowTemplate = tempChildRowTemplate.Replace("$StepNo$", stpnmb);
                // update comments
                if (string.IsNullOrWhiteSpace(com) == false)
                {
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$Child-Steps$", com);
                }
                else
                    tempChildRowTemplate = tempChildRowTemplate.Replace("$Child-Steps$", stpdesp);

                string htmlAttach = string.Empty;
                string htmlAttachPath = string.Empty;

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
                    Common.Property.TotalStepFail++;
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
                    Common.Property.TotalStepWarning++;
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
                    Common.Property.TotalStepPass++;
                }
                sbChildHtmlRows.Append(tempChildRowTemplate);
                rowNo++;
            }
            return sbChildHtmlRows.ToString();
        }
        static void CreateZip(string[] xmlFilesLocation)
        {
            try
            {
                using (ZipFile zip = new ZipFile())
                {
                    foreach (string xmlLocation in xmlFilesLocation)
                    {
                        // Path to directory of files to compress and decompress.                               
                        string directory = Path.GetDirectoryName(xmlLocation);
                        if (Directory.Exists(directory))
                        {
                            FileInfo fileInfo = new FileInfo(xmlLocation);
                            zip.AddDirectory(directory, fileInfo.Directory.Parent.Name + "/" + fileInfo.Directory.Name);
                        }
                    }
                    zip.AddFile(Common.Property.HtmlFileLocation + "\\HtmlReport.html", "./");
                    if (File.Exists(Common.Property.HtmlFileLocation + "\\" + Property.ReportZipFileName))
                        //delete zip file if already exists
                        File.Delete(Common.Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);
                    zip.Save(Common.Property.HtmlFileLocation + "\\" + Property.ReportZipFileName);
                }
            }
            catch (Exception ex1)
            {
                throw ex1;
            }
        }
    }
}
