/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Program.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Read the Parameters for the Reporting 
*****************************************************************************/
using System;
using System.IO;
using System.Linq;
using Common;
using Reporting;
using System.Windows.Forms;

namespace KRYPTONCreateReport
{
    class Program
    {
        public static string XmlFileName = string.Empty;
        public static string HtmlFileName = string.Empty;
        public static string Environment = string.Empty;
        public static string EmailFor = "end";
        public static string TestSuite = string.Empty;
        public static string EmailRequired = "no";
        public static string CurrentProjectpath = string.Empty;
        public static string XmlrootFolderPath = string.Empty;


        static void Main(string[] args)
        {

            //setting the application path
            string applicationPath = Application.ExecutablePath;
            int applicationFilePath = applicationPath.LastIndexOf("\\", StringComparison.Ordinal);

            if (applicationFilePath >= 0)
            {
                Property.ApplicationPath = applicationPath.Substring(0, applicationFilePath + 1);
            }
            else
            {
                DirectoryInfo dr = new DirectoryInfo("./");
                Property.ApplicationPath = dr.FullName;
            }

            try
            {

                using (StreamReader sr = new StreamReader(Path.Combine(Property.ApplicationPath, "root.ini")))
                {
                    string currentProjectName;
                    while ((currentProjectName = sr.ReadLine()) != null)
                    {
                        if (currentProjectName.ToLower().Contains("projectpath"))
                        {
                            CurrentProjectpath = currentProjectName.Substring(currentProjectName.IndexOf(':') + 1).Trim();
                            break;
                        }
                    }
                }
                if (!Path.IsPathRooted(CurrentProjectpath))
                    CurrentProjectpath = Path.Combine(Property.ApplicationPath, CurrentProjectpath);
                Property.IniPath = Path.GetFullPath(CurrentProjectpath);
                if (!Directory.Exists(Property.IniPath))
                    Property.IniPath = Property.ApplicationPath;

            }
            catch (Exception)
            {
                Property.IniPath = Property.ApplicationPath;
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
                if (File.Exists(Property.ApplicationPath + "xmlFiles.txt"))
                {
                    StreamReader xmlTextFileRead = new StreamReader(Property.ApplicationPath + "xmlFiles.txt");
                    XmlFileName = xmlTextFileRead.ReadToEnd();
                    xmlTextFileRead.Close();
                }
            }
            catch
            {
                // ignored
            }
            Environment = Utility.GetParameter("Environment");
            ReadCommandLineArguments(args);
            Utility.SetParameter(Property.Environment, Environment);
            XmlFileName = XmlFileName.Replace("\n", ",");
            XmlFileName = XmlFileName.Replace("\r", string.Empty);
            XmlFileName = XmlFileName.Replace("\t", string.Empty);


            string htmlFile = string.Empty;
            bool isSummaryRequiredinResultsFolder = false;
            bool callFromKryptonVbScriptGrid = true;
            if (XmlrootFolderPath.Length > 0)
            {
                isSummaryRequiredinResultsFolder = Utility.GetParameter("SummaryReportRequired").ToLower().Equals("true");
                callFromKryptonVbScriptGrid = false;
            }
            if (!EmailFor.Equals("start", StringComparison.OrdinalIgnoreCase))
            {
                if (XmlrootFolderPath.Length > 0)
                {
                    LogFile.AllXmlFilesLocation = GetFileNames(XmlrootFolderPath);
                    Property.HtmlFileLocation = Path.Combine("", XmlrootFolderPath);

                }
                else
                {

                    LogFile.AllXmlFilesLocation = string.Empty;

                    string[] xmlFiles = XmlFileName.Split(',');

                    for (int cntXml = 0; cntXml < xmlFiles.Length; cntXml++)
                    {
                        if (string.IsNullOrWhiteSpace(LogFile.AllXmlFilesLocation))
                            LogFile.AllXmlFilesLocation = xmlFiles[cntXml];
                        else
                            LogFile.AllXmlFilesLocation = LogFile.AllXmlFilesLocation + ";" + xmlFiles[cntXml];
                    }
                    try
                    {
                        var directoryInfo = new FileInfo(xmlFiles[0]).Directory;
                        if (directoryInfo != null)
                            if (directoryInfo.Parent != null)
                                if (directoryInfo.Parent.Parent != null)
                                    Property.HtmlFileLocation = directoryInfo.Parent.Parent.FullName;
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
                if (!string.IsNullOrWhiteSpace(LogFile.AllXmlFilesLocation))
                {
                    Property.Date_Time = Utility.GetParameter("DateTimeFormat").Replace("/", "\\/");
                    htmlFile = "HtmlReport-" + HtmlFileName + "-" + DateTime.Now.ToString("ddMMyyhhmmss") +
                                        ".html";
                    Property.CompanyLogo = Path.IsPathRooted(Utility.GetParameter("CompanyLogo")) ? Utility.GetParameter("CompanyLogo") : string.Concat(Property.IniPath, Utility.GetParameter("CompanyLogo"));
                    if (!File.Exists(Property.CompanyLogo))
                        Property.CompanyLogo = string.Empty;
                    LogFile.CreateHtmlReport(htmlFile, false, false, Property.IsSauceLabExecution);
                    //Always create summpary report because it is called from grid and this summary report is used in grid to send email.
                    HtmlReport.CreateHtmlReport(htmlFile, false, false, Property.IsSauceLabExecution, isSummaryRequiredinResultsFolder, true, callFromKryptonVbScriptGrid, true);
                }
                htmlFile = htmlFile.Replace(".html", "smail.html"); // to send mail
                htmlFile = Path.Combine(Property.HtmlFileLocation, htmlFile);
                if (XmlrootFolderPath.Length > 0)
                    EmailRequired = (Utility.GetParameter("EmailNotification").Equals("true", StringComparison.OrdinalIgnoreCase)) ? "yes" : "no";

                if (string.Equals(EmailRequired, "yes"))
                {
                    try
                    {
                        Utility.EmailNotification(EmailFor, false, htmlFile);
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
                        if (!callFromKryptonVbScriptGrid)
                            File.Delete(htmlFile);
                    }
                    catch
                    {
                        // ignored
                    }
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
                    if (arguments[1].Trim().IndexOf(".txt", StringComparison.Ordinal) < 0)
                    {
                        XmlFileName = arguments[1].Trim();
                    }
                    else
                    {
                        try
                        {
                            StreamReader xmlTextFileRead = new StreamReader(arguments[1].Trim());
                            XmlFileName = xmlTextFileRead.ReadToEnd();
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
                    Environment = arguments[1];
                }
                if (arguments[0].ToLower().Trim().Equals("xmlrootfolderpath", StringComparison.OrdinalIgnoreCase))
                {
                    XmlrootFolderPath = arguments[1].Trim(); 
                }
                if (arguments[0].Trim().Equals("htmlFileName", StringComparison.OrdinalIgnoreCase))
                {
                    HtmlFileName = arguments[1];
                }
                if (arguments[0].Trim().Equals("emailFor", StringComparison.OrdinalIgnoreCase))
                {
                    EmailFor = arguments[1];
                }
                if (arguments[0].Trim().Equals("testSuite", StringComparison.OrdinalIgnoreCase))
                {
                    TestSuite = arguments[1];
                    Utility.SetParameter("TestSuite", TestSuite);
                    Utility.SetParameter("TestCaseId", string.Empty);
                }
                if (arguments[0].Trim().Equals("emailRequired", StringComparison.OrdinalIgnoreCase))
                {
                    EmailRequired = arguments[1];

                }
            }
        }

        public static void ReadIniFiles()
        {
            FileInfo reportSettingsFileInfo = new FileInfo(Path.Combine(Property.ApplicationPath, Property.ReportSettingsFile));
            using(StreamReader reportSettingsReader = new StreamReader(reportSettingsFileInfo.FullName))
            Utility.StoreReaderContent(reportSettingsReader);

            FileInfo paramSettingsFileInfo = new FileInfo(Path.Combine(Property.IniPath, Property.ParameterFileName));
            using(StreamReader paramSettingsReader = new StreamReader(paramSettingsFileInfo.FullName))
            Utility.StoreReaderContent(paramSettingsReader);

            FileInfo emailFileInfo = new FileInfo(Path.Combine(Property.IniPath, Property.EmailNotificationFile));
            using(StreamReader emailReader = new StreamReader(emailFileInfo.FullName))
            Utility.StoreReaderContent(emailReader);
            Property.EmailStartTemplate = Path.IsPathRooted(Utility.GetParameter("EmailStartTemplate")) ? Utility.GetParameter("EmailStartTemplate") : string.Concat(Property.IniPath, Utility.GetParameter("EmailStartTemplate"));
            Property.EmailEndTemplate = Path.IsPathRooted(Utility.GetParameter("EmailEndTemplate")) ? Utility.GetParameter("EmailEndTemplate") : string.Concat(Property.IniPath, Utility.GetParameter("EmailEndTemplate"));
        }

        private static string GetFileNames(string xmlPath)
        {
            string combinedPath = string.Empty;
            if (Directory.Exists(xmlPath))
            {
                string[] xmlfiles = Directory.GetFiles(xmlPath, "*.xml", SearchOption.AllDirectories);
                combinedPath = xmlfiles.Aggregate(combinedPath, (current, xmlfile) => current + ";" + xmlfile);
            }
            return combinedPath;
        }

    }
}

