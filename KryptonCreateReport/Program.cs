/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Program.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Read the Parameters for the Reporting 
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using Common;
using Reporting;
using PreMailer.Net;
using System.Windows.Forms;

namespace KRYPTONCreateReport
{
    class Program
    {
        public static string xmlFileName = string.Empty;
        public static string htmlFileName = string.Empty;
        public static string environment = string.Empty;
        public static string emailFor = "end";
        public static string testSuite = string.Empty;
        public static string emailRequired = "no";
        public static string currentProjectpath = string.Empty;
        public static string xmlrootFolderPath = string.Empty;


        static void Main(string[] args)
        {

            //setting the application path
            string applicationPath = Application.ExecutablePath;
            int applicationFilePath = applicationPath.LastIndexOf("\\");

            if (applicationFilePath >= 0)
            {
                Common.Property.ApplicationPath = applicationPath.Substring(0, applicationFilePath + 1);
            }
            else
            {
                DirectoryInfo dr = new DirectoryInfo("./");
                Common.Property.ApplicationPath = dr.FullName;
            }

            try
            {

                using (StreamReader sr = new StreamReader(Path.Combine(Common.Property.ApplicationPath, "root.ini")))
                {
                    string currentProjectName = string.Empty;
                    while ((currentProjectName = sr.ReadLine()) != null)
                    {
                        if (currentProjectName.ToLower().Contains("projectpath"))
                        {
                            currentProjectpath = currentProjectName.Substring(currentProjectName.IndexOf(':') + 1).Trim();
                            break;
                        }
                    }
                }
                if (!Path.IsPathRooted(currentProjectpath))
                    currentProjectpath = Path.Combine(Common.Property.ApplicationPath, currentProjectpath);
                Common.Property.IniPath = Path.GetFullPath(currentProjectpath);
                if (!Directory.Exists(Common.Property.IniPath))
                    Common.Property.IniPath = Common.Property.ApplicationPath;

            }
            catch (Exception es)
            {
                Common.Property.IniPath = Common.Property.ApplicationPath;
            }


            try
            {
                ReadIniFiles();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            try
            {
                if (File.Exists(Common.Property.ApplicationPath + "xmlFiles.txt"))
                {
                    StreamReader xmlTextFileRead = new StreamReader(Common.Property.ApplicationPath + "xmlFiles.txt");
                    xmlFileName = xmlTextFileRead.ReadToEnd();
                    xmlTextFileRead.Close();
                }
            }
            catch
            {

            }
            environment = Common.Utility.GetParameter("Environment");
            ReadCommandLineArguments(args);
            Common.Utility.SetParameter(Common.Property.ENVIRONMENT, environment);
            xmlFileName = xmlFileName.Replace("\n", ",");
            xmlFileName = xmlFileName.Replace("\r", string.Empty);
            xmlFileName = xmlFileName.Replace("\t", string.Empty);


            string htmlFile = string.Empty;
            bool IsSummaryRequiredinResultsFolder = false;
            bool CallFromKryptonVBScriptGrid = true;
            if (xmlrootFolderPath.Length > 0)
            {
                IsSummaryRequiredinResultsFolder = Utility.GetParameter("SummaryReportRequired").ToLower().Equals("true");
                CallFromKryptonVBScriptGrid = false;
            }
            if (!emailFor.Equals("start", StringComparison.OrdinalIgnoreCase))
            {
                if (xmlrootFolderPath.Length > 0)
                {
                    LogFile.allXmlFilesLocation = GetFileNames(xmlrootFolderPath);
                    Property.HtmlFileLocation = Path.Combine("", xmlrootFolderPath);

                }
                else
                {

                    LogFile.allXmlFilesLocation = string.Empty;

                    string[] xmlFiles = xmlFileName.Split(',');

                    for (int cntXml = 0; cntXml < xmlFiles.Length; cntXml++)
                    {
                        if (string.IsNullOrWhiteSpace(LogFile.allXmlFilesLocation))
                            LogFile.allXmlFilesLocation = xmlFiles[cntXml];
                        else
                            LogFile.allXmlFilesLocation = LogFile.allXmlFilesLocation + ";" + xmlFiles[cntXml];
                    }
                    try
                    {
                        Common.Property.HtmlFileLocation = new FileInfo(xmlFiles[0]).Directory.Parent.Parent.FullName;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error in path");
                    }
                }
            }
            else
            {
                Property.ExecutionStartDateTime = DateTime.Now.ToString(Utility.GetParameter("DateTimeFormat"));
                Property.ExecutionEndDateTime = DateTime.Now.ToString(Utility.GetParameter("DateTimeFormat"));
            }

            try
            {
                htmlFile = string.Empty;
                if (!string.IsNullOrWhiteSpace(LogFile.allXmlFilesLocation))
                {
                    Common.Property.DATE_TIME = Common.Utility.GetParameter("DateTimeFormat").Replace("/", "\\/");
                    htmlFile = "HtmlReport-" + htmlFileName + "-" + DateTime.Now.ToString("ddMMyyhhmmss") +
                                        ".html";
                    if (Path.IsPathRooted(Common.Utility.GetParameter("CompanyLogo")))
                        Common.Property.CompanyLogo = Common.Utility.GetParameter("CompanyLogo");
                    else
                        Common.Property.CompanyLogo = string.Concat(Common.Property.IniPath, Common.Utility.GetParameter("CompanyLogo"));
                    if (!File.Exists(Common.Property.CompanyLogo))
                        Common.Property.CompanyLogo = string.Empty;
                    LogFile.CreateHtmlReport(htmlFile, false, false, Property.isSauceLabExecution);
                    //Always create summpary report because it is called from grid and this summary report is used in grid to send email.
                    Reporting.HTMLReport.CreateHtmlReport(htmlFile, false, false, Property.isSauceLabExecution, IsSummaryRequiredinResultsFolder, true, CallFromKryptonVBScriptGrid, true);
                }
                htmlFile = htmlFile.Replace(".html", "smail.html"); // to send mail
                htmlFile = Path.Combine(Property.HtmlFileLocation, htmlFile);
                if (xmlrootFolderPath.Length > 0)
                    emailRequired = (Utility.GetParameter("EmailNotification").Equals("true", StringComparison.OrdinalIgnoreCase)) ? "yes" : "no";

                if (string.Equals(emailRequired, "yes"))
                {
                    try
                    {
                        Common.Utility.EmailNotification(emailFor, false, htmlFile);

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);

                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);


            }
            finally
            {
                if (File.Exists(htmlFile)) //delete mailhtml
                {
                    try
                    {
                        if (!CallFromKryptonVBScriptGrid)
                            File.Delete(htmlFile);
                    }
                    catch { }

                }
            }
        }

        private static void ReadCommandLineArguments(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                string[] arguments = args[i].Split('=');
                if (arguments[0].Trim().Equals("xmlFiles", StringComparison.OrdinalIgnoreCase))
                {
                    if (arguments[1].Trim().IndexOf(".txt") < 0)
                    {
                        xmlFileName = arguments[1].Trim();
                    }
                    else
                    {
                        try
                        {
                            StreamReader xmlTextFileRead = new StreamReader(arguments[1].Trim());
                            xmlFileName = xmlTextFileRead.ReadToEnd();
                            xmlTextFileRead.Close();
                        }
                        catch (Exception e)
                        {

                            Console.WriteLine(e.Message);
                        }
                    }
                }
                if (arguments[0].Trim().Equals("environment", StringComparison.OrdinalIgnoreCase))
                {
                    environment = arguments[1];
                }
                if (arguments[0].ToLower().Trim().Equals("xmlrootfolderpath", StringComparison.OrdinalIgnoreCase))
                {
                    xmlrootFolderPath = arguments[1].Trim(); 
                }
                if (arguments[0].Trim().Equals("htmlFileName", StringComparison.OrdinalIgnoreCase))
                {
                    htmlFileName = arguments[1];
                }
                if (arguments[0].Trim().Equals("emailFor", StringComparison.OrdinalIgnoreCase))
                {
                    emailFor = arguments[1];
                }
                if (arguments[0].Trim().Equals("testSuite", StringComparison.OrdinalIgnoreCase))
                {
                    testSuite = arguments[1];
                    Common.Utility.SetParameter("TestSuite", testSuite);
                    Common.Utility.SetParameter("TestCaseId", string.Empty);
                }
                if (arguments[0].Trim().Equals("emailRequired", StringComparison.OrdinalIgnoreCase))
                {
                    emailRequired = arguments[1];

                }
            }
        }

        public static void ReadIniFiles()
        {
            FileInfo reportSettingsFileInfo = new FileInfo(Path.Combine(Common.Property.ApplicationPath, Property.ReportSettingsFile));
            StreamReader reportSettingsReader = new StreamReader(reportSettingsFileInfo.FullName);
            Common.Utility.StoreReaderContent(reportSettingsReader);

            FileInfo paramSettingsFileInfo = new FileInfo(Path.Combine(Common.Property.IniPath, Property.ParameterFileName));
            StreamReader paramSettingsReader = new StreamReader(paramSettingsFileInfo.FullName);
            Common.Utility.StoreReaderContent(paramSettingsReader);

            FileInfo emailFileInfo = new FileInfo(Path.Combine(Common.Property.IniPath, Property.EmailNotificationFile));
            StreamReader emailReader = new StreamReader(emailFileInfo.FullName);
            Common.Utility.StoreReaderContent(emailReader);


            if (Path.IsPathRooted(Common.Utility.GetParameter("EmailStartTemplate")))
                Common.Property.EmailStartTemplate = Common.Utility.GetParameter("EmailStartTemplate");
            else
                Common.Property.EmailStartTemplate = string.Concat(Common.Property.IniPath, Common.Utility.GetParameter("EmailStartTemplate"));

            if (Path.IsPathRooted(Common.Utility.GetParameter("EmailEndTemplate")))
                Common.Property.EmailEndTemplate = Common.Utility.GetParameter("EmailEndTemplate");
            else
                Common.Property.EmailEndTemplate = string.Concat(Common.Property.IniPath, Common.Utility.GetParameter("EmailEndTemplate"));
        }

        private static string GetFileNames(string xmlPath)
        {
            string combinedPath = string.Empty;
            if (Directory.Exists(xmlPath))
            {
                string[] xmlfiles = Directory.GetFiles(xmlPath, "*.xml", SearchOption.AllDirectories);
                foreach (string xmlfile in xmlfiles)
                    combinedPath = combinedPath + ";" + xmlfile;
            }
            return combinedPath;
        }

    }
}

