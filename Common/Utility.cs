/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: Common.Utility.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: All Utility Functionalities goes here.
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text;
using System.IO;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Threading;
using System.Data;
using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;


namespace Common
{
    public class Utility
    {

        public static Dictionary<int, string> driverKeydic = new Dictionary<int, string>();

        /// <summary>
        /// Match the expected and actual values based on options specified.
        /// </summary>
        /// <param name="expectedValue"> string expected value to match.</param>
        /// <param name="actualValue"> string actual value to match.</param>
        /// <returns>bool result of match.</returns>
        public static bool doKeywordMatch(string expectedValue, string actualValue)
        {

            if (driverKeydic.ContainsValue("ignorespace"))
            {
                expectedValue = Regex.Replace(expectedValue, @"[\s]", String.Empty).ToLower();
                actualValue = Regex.Replace(actualValue, @"[\s]", String.Empty).ToLower();
            }
            if (driverKeydic.ContainsValue("ignorecase"))
            {
                expectedValue = expectedValue.ToLower();
                actualValue = actualValue.ToLower();
            }
            if (!driverKeydic.ContainsValue("partialmatch") && !driverKeydic.ContainsValue("partmatch") && driverKeydic.ContainsValue("exactmatch"))
                return actualValue.Equals(expectedValue);
            else
                return actualValue.Contains(expectedValue);
        }

        /// <summary>
        ///  Generates a unique character string for the specified length
        /// </summary>
        /// <param name="CharCount">Length of string to be generated,by default it is 10.</param>
        /// <returns>Unique string with specified length.</returns>
        public static string GenerateUniqueString(int CharCount = 10)
        {
            Random random = new Random();
            int randomnum;
            string randomString = string.Empty;
            for (int i = 0; i < CharCount; i++)
            {
                randomnum = random.Next(97, 123);
                randomString = randomString + (char)randomnum;
            }
            return randomString;
        }

        /// <summary>
        /// Generate a numeral string such as zip code or mobile number
        /// </summary>
        /// <param name="CharCount"> Length of number to be generated,by default it is 5,</param>
        /// <returns>Unique number with specified length in string format.</returns>
        public static string GenerateUniqueNumeral(int CharCount = 5)
        {
            Random random = new Random();
            int randomnum;
            string randomString = string.Empty;
            for (int i = 0; i < CharCount; i++)
            {
                randomnum = random.Next(1, 9);
                randomString = randomString + randomnum;
            }
            return randomString;
        }

        public static string GetOSVersion()
        {
            string osName = "";
            try
            {
                string s = System.Environment.OSVersion.ToString();
                if (s.Contains("5.1"))
                    osName = "Windows XP";
                else if (s.Contains("5.2"))
                    osName = "Windows 2003";
                else if (s.Contains("5.0"))
                    osName = "Windows 2000";
                else if (s.Contains("4.0"))
                    osName = "Windows 2000";
                else if (s.Contains("6.0"))
                    osName = "Windows 2000";
                else
                    osName = "Unknown - probably Win 9x";
            }
            catch (Exception e)
            {
                throw new KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0039").Replace("{MSG}", e.Message));
            }
            return osName;
        }

        /// <summary>
        /// Stores a variable in run time dictionary. Variable name in run time dictionary are always lower case.
        ///This will also overwrite any previously written variable value.
        /// </summary>
        /// <param name="varName"></param>
        /// <param name="varValue"></param>
        public static void SetVariable(string varName, string varValue)
        {
            varName = varName.ToLower();
            if (Common.Property.runtimedic.ContainsKey(varName))
                Common.Property.runtimedic[varName] = varValue;
            else
                Common.Property.runtimedic.Add(varName, varValue);
        }


        /// <summary>
        /// Retrives value of variable from run time dictionary. Return ZERO length string if variable
        ///             could not be found in dictionary.
        /// </summary>
        /// <param name="varName"></param>
        /// <returns></returns>
        public static string GetVariable(string varName)
        {
            varName = varName.ToLower();
            if (!Common.Property.runtimedic.ContainsKey(varName))
                return varName;

            if (string.IsNullOrEmpty(Property.runtimedic[varName]))
                return Property.runtimedic[varName];


            if (varName == "firefoxprofilepath")
            {
                if (!Path.IsPathRooted(Common.Property.runtimedic[varName]))
                {
                    Common.Property.runtimedic[varName] = string.Concat(Common.Property.IniPath, Common.Property.runtimedic[varName]);
                    Common.Property.runtimedic[varName] = Path.GetFullPath(Common.Property.runtimedic[varName]);
                }
            }
            return Common.Property.runtimedic[varName];
        }

        /// <summary>
        /// Get the parameter from parameter dictionary.
        /// </summary>
        /// <param name="parameterName">Key String for which values needed.</param>
        /// <returns>String value for given parameter.</returns>
        public static string GetParameter(string parameterName)
        {

            parameterName = parameterName.ToLower();
            if (!Common.Property.parameterdic.ContainsKey(parameterName))
                return parameterName;
            if (string.IsNullOrEmpty(Property.parameterdic[parameterName]))
                return Property.runtimedic[parameterName];
            if (parameterName == "firefoxprofilepath")
            {
                if (!Path.IsPathRooted(Common.Property.parameterdic[parameterName]))
                {
                    Common.Property.parameterdic[parameterName] = string.Concat(Common.Property.IniPath, Common.Property.parameterdic[parameterName]);
                    Common.Property.parameterdic[parameterName] = Path.GetFullPath(Common.Property.parameterdic[parameterName]);
                }
            }
            return Common.Property.parameterdic[parameterName];

        }

        /// <summary>
        /// SetParameter to Parameter dictionary.
        /// </summary>
        /// <param name="paramName">Parameter String for which value to be set.</param>
        /// <param name="paramValue">Parameter Value which will be set.</param>
        public static void SetParameter(string paramName, string paramValue)
        {
            paramName = paramName.ToLower();
            if (Common.Property.parameterdic.ContainsKey(paramName))
                Common.Property.parameterdic[paramName] = paramValue;
            else
                Common.Property.parameterdic.Add(paramName, paramValue);
        }

        /// <summary>
        /// Return all processes started.
        /// </summary>
        /// <returns>string array of processes.</returns>
        public static string[] getAllProcesses()
        {
            string[] processes = new string[Common.Property.processLists.Count];
            for (int i = 0; i < Common.Property.processLists.Count; i++)
            {
                processes[i] = Common.Property.processLists[i].ToString();
            }
            return processes;
        }

        /// <summary>
        /// Set process string top global arraylist.
        /// </summary>
        /// <param name="processName">string process name to add.</param>
        public static void setProcessParameter(string processName)
        {
            Common.Property.processLists.Add(processName);
        }

        /// <summary>
        /// Replaces variables in a string. Variables are used with a convention of {$varName}.
        ///             Variable is not replaced if its values could not be found.
        ///Replace variable by {$varName} after discussion
        /// </summary>
        /// <param name="inString">Input string, in which variable needs to be replaced.</param>
        /// <returns>String with variables replaced</returns>
        public static string ReplaceVariablesInString(string inString)
        {
            int stindex;
            int endindex;
            stindex = 0;
            endindex = 0;
            for (int v = 0; ; v++)
            {
                if (inString.Contains("{$"))
                {

                    stindex = inString.IndexOf("{$", endindex);
                    //Break if no more variables found the string
                    if (stindex < 0)
                    {
                        break;
                    }

                    endindex = inString.IndexOf("}", stindex);
                    //Break if end } is not there in the string
                    if (endindex < 0)
                    {
                        break;
                    }
                    string KeyVariable = inString.Substring(stindex + 2, (endindex - stindex - 2));

                    //Retrieve value of variable from run time dictionary
                    string value = GetVariable(KeyVariable);

                    //If variable name and returned value are same, consider no variable found, and do not replace original string
                    if (value.ToLower().Equals(KeyVariable.ToLower()) == false)
                    {
                        inString = inString.Replace("{$" + KeyVariable + "}", value);
                    }
                    endindex = stindex + 1;
                }
                else
                {
                    break;
                }
            }
            return inString;
        }

        /// <summary>
        /// Collect data from Parameters.ini file and store in a dictionary called parameterdic.
        /// parameterdic is a global dictiomary that contains all the parameters needed to framework.
        /// </summary>
        public static void CollectkeyValuePairs()
        {
            StreamReader reader = null;
            try
            {
                FileInfo parameterFileInfo = new FileInfo(Path.Combine(Common.Property.IniPath, Common.Property.ParameterFileName));
                reader = new StreamReader(parameterFileInfo.FullName);
                StoreReaderContent(reader);

                // Retreiving the content of report settings in INI file
                try
                {
                    FileInfo reportSettingsFileInfo = new FileInfo(Common.Property.ApplicationPath + Property.ReportSettingsFile);
                    StreamReader reportSettingsReader = new StreamReader(reportSettingsFileInfo.FullName);
                    StoreReaderContent(reportSettingsReader);
                }
                catch (Exception e)
                {
                    throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0042"), e.Message);
                }

                // Retreiving the contents of SauceLabs.ini file
                if (Property.parameterdic[Property.BrowserString].ToLower() == Property.SAUCELABS.ToLower())
                {
                    try
                    {
                        if (File.Exists(Path.Combine(Common.Property.IniPath, Property.SAUCELABS_PARAMETER_FILE)))
                        {
                            FileInfo SauceLabsFileInfo = new FileInfo(Path.Combine(Common.Property.IniPath, Property.SAUCELABS_PARAMETER_FILE));
                            StreamReader SauceLabsReader = new StreamReader(SauceLabsFileInfo.FullName);
                            StoreReaderContent(SauceLabsReader);
                        }
                        else
                            throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0037"), "saucelabs.ini");
                    }
                    catch (Exception e)
                    {

                        throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0059"), e.Message);
                    }
                }
                //END of sauce file access code
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0044"), e.Message);
            }
        }


        /// <summary>
        /// Collect data from other ini such as QA and Test Manager
        /// </summary>
        public static void UpdateKeyValuePairs()
        {
            try
            {
                // Retreiving the content of database specific INI files,based on the key/value mentioned
                // in Parameter.ini file
                try
                {
                    FileInfo databaseFileInfo = new FileInfo(Path.Combine(Common.Property.EnvironmentFileLocation, GetParameter(Property.ENVIRONMENT) + ".ini"));
                    StreamReader dbreader = new StreamReader(databaseFileInfo.FullName);
                    StoreReaderContent(dbreader);
                }
                catch (Exception e)
                {
                    throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0041"), e.Message);
                }

                // Retreiving the content of test manager specific ini file
                // This file will be named as QC.ini, XStudio.ini or MSTestManager.ini etc
                {
                    FileInfo databaseFileInfo = new FileInfo(Path.Combine(Common.Property.ApplicationPath, GetParameter("ManagerType") + ".ini"));
                    if (File.Exists(databaseFileInfo.FullName))
                    {
                        StreamReader dbreader = new StreamReader(databaseFileInfo.FullName);
                        StoreReaderContent(dbreader);
                    }
                }
                //Parsing any other extra ini file, that can be passed as iniFiles=TestManager.ini, Custom.ini etc.
                //This section is highly powerfull, using which you can pass on any ini file to be parsed
                string extraIniFiles = Common.Utility.GetVariable("iniFiles");
                string[] arrIniFiles = extraIniFiles.Split(',');
                foreach (string iniFile in arrIniFiles)
                {
                    string iniFileLocation = Path.Combine(Common.Property.ApplicationPath, iniFile.Trim());
                    if (File.Exists(iniFileLocation))
                    {
                        FileInfo databaseFileInfo = new FileInfo(iniFileLocation);
                        StreamReader dbreader = new StreamReader(databaseFileInfo.FullName);
                        StoreReaderContent(dbreader);
                    }
                }
            }
            catch (Exception e)
            {
                throw new Common.KryptonException(Utility.GetCommonMsgVariable("KRYPTONERRCODE0044"), e.Message);
            }
        }
        /// <summary>
        /// Generate string (in executable arguments format ) from specified Dictionary.
        /// </summary>
        /// <param name="dicToProcess">Dictionary to process ie. Dictionary<string,string></param>
        /// <returns>String</returns>
        public static string generateDictionaryContentInArgumentsFormat(Dictionary<string, string> dicToProcess)
        {
            string argumentString = string.Empty;
            try
            {
                foreach (KeyValuePair<string, string> keyValuePair in dicToProcess)
                {
                    if (argumentString.Equals(string.Empty)) { argumentString = argumentString + "\"" + keyValuePair.Key + "=" + keyValuePair.Value + "\""; }
                    else argumentString = argumentString + " " + "\"" + keyValuePair.Key + "=" + keyValuePair.Value + "\"";
                }


            }
            catch (Exception) { }
            return argumentString;
        }

        public static string createTempINIfile(Dictionary<string, string> dic, Dictionary<string, string> otherDic)
        {
            string tempFilePAth = string.Empty;

            //This will avoid sinking of dictionaries.
            Dictionary<string, string> tempDIC = new Dictionary<string, string>(dic);

            try
            {
                foreach (KeyValuePair<string, string> KVPair in otherDic)
                {
                    if (dic.ContainsKey(KVPair.Key)) { tempDIC[KVPair.Key] = otherDic[KVPair.Key]; }
                    else { tempDIC.Add(KVPair.Key.ToLower(), KVPair.Value); }
                }
                tempFilePAth = Property.ApplicationPath + "//" + Property.TEMPINIFILENAME + ".ini";
                FileStream stream = File.Create(tempFilePAth);
                StreamWriter writer = new StreamWriter(stream, Encoding.UTF8);
                // for merging mobile build
                foreach (KeyValuePair<string, string> keyValuePair in tempDIC)
                {
                    writer.WriteLine(keyValuePair.Key + " : " + keyValuePair.Value);
                }
                writer.Dispose();
                stream.Dispose();
            }
            catch (Exception e)
            {
                // Console.WriteLine(e.Message);
            }
            return tempFilePAth;
        }

        /// <summary>
        ///Verify the presence of all the parameters in parameters.ini file
        ///It also checks for the values of the mandatory parameters.
        /// </summary>
        /// <param name="reader">StreamReader object of parameters.ini</param>
        public static void ValidateParameters()
        {
            bool validation = true;
            //    testCaseOrSuite;
            string errorMessage = string.Empty;
            Dictionary<String, int> allParameters = new Dictionary<string, int>();

            //Please make an entry if added a new parameter in parameters.ini file
            allParameters.Add("browser", 0);
            allParameters.Add("driver", 0);
            allParameters.Add("testcaselocation", 1);
            allParameters.Add("testdatalocation", 1);
            allParameters.Add("testcasefileextension", 0);
            allParameters.Add("reusabledefaultfile", 0);
            allParameters.Add("orlocation", 1);
            allParameters.Add("orfilename", 1);
            allParameters.Add("recoverfrompopuplocation", 1);
            allParameters.Add("recoverfrompopupfilename", 1);
            allParameters.Add("recoverfrombrowserlocation", 1);
            allParameters.Add("recoverfrombrowserfilename", 1);
            allParameters.Add("dbtestdatalocation", 1);
            allParameters.Add("logdestinationfolder", 1);
            allParameters.Add("testcaseid", 2); // verified in conjunction
            allParameters.Add("testsuite", 2);  // verified in conjunction
            allParameters.Add("managertype", 0);
            allParameters.Add("timeformat", 0);
            allParameters.Add("dateformat", 0);
            allParameters.Add("datetimeformat", 0);
            allParameters.Add("companylogo", 0);
            allParameters.Add("failedcountforexit", 0);
            allParameters.Add("endexecutionwaitrequired", 0);
            allParameters.Add("dbconnectionstring", 1);
            allParameters.Add("defaultdb", 1);
            allParameters.Add("dbserver", 1);
            allParameters.Add("dbqueryfilepath", 1);
            allParameters.Add("debugmode", 0);
            allParameters.Add("errorcaptureas", 0);
            allParameters.Add("environment", 0);
            allParameters.Add("environmentsetupbatch", 0);
            allParameters.Add("testcaseidseperator", 0);
            allParameters.Add("testcaseidparameter", 0);
            allParameters.Add("snapshotoption", 0);
            allParameters.Add("runremoteexecution", 0);
            allParameters.Add("runonremotebrowserurl", 0);
            allParameters.Add("emailnotification", 0);
            allParameters.Add("emailnotificationfrom", 0);
            allParameters.Add("emailsmtpserver", 0);
            allParameters.Add("emailsmtpport", 0);
            allParameters.Add("emailsmtpusername", 0);
            allParameters.Add("emailsmtppassword", 0);
            allParameters.Add("emailstarttemplate", 1);
            allParameters.Add("emailendtemplate", 1);
            allParameters.Add("keepreporthistory", 0);
            allParameters.Add("validatesetup", 0);
            allParameters.Add("objecttimeout", 0);
            allParameters.Add("globaltimeout", 0);
            allParameters.Add("testmode", 0);
            allParameters.Add("closebrowseroncompletion", 0);
            allParameters.Add("firefoxprofilepath", 0);
            allParameters.Add("addonspath", 0);
            allParameters.Add("applicationurl", 1);
            allParameters.Add("maxtimeoutforpageload", 0);
            allParameters.Add("mintimeoutforpageload", 0);
            allParameters.Add("scriptlanguage", 0);
            allParameters.Add("recoverycount", 0);
            allParameters.Add("parallelrecoverysheetname", 1);
            allParameters.Add("environmentfilelocation", 1);
            allParameters.Add("startparallelrecovery", 2);
            // allParameters.Add("Waitforalert", 0);
            try
            {
                for (int i = 0; i < allParameters.Count; i++)
                {
                    if (Property.parameterdic.ContainsKey(allParameters.ElementAt(i).Key) == false)
                    {
                        validation = false;
                        errorMessage = Utility.GetCommonMsgVariable("KRYPTONERRCODE0045").Replace("{MSG}", allParameters.ElementAt(i).Key);
                        break;
                    }

                    if (allParameters.ElementAt(i).Value == 0) continue;
                    else if (allParameters.ElementAt(i).Value == 1)
                    {
                        if (string.IsNullOrWhiteSpace(GetParameter(allParameters.ElementAt(i).Key)))
                        {
                            validation = false;
                            errorMessage = Utility.GetCommonMsgVariable("KRYPTONERRCODE0046").Replace("{MSG}", allParameters.ElementAt(i).Key);
                            break;
                        }
                    }
                    else if (allParameters.ElementAt(i).Value == 2)
                    {
                        if (string.IsNullOrWhiteSpace(GetParameter("TestCaseId")) && string.IsNullOrWhiteSpace(GetParameter("TestSuite")))
                        {
                            validation = false;
                            errorMessage = Utility.GetCommonMsgVariable("KRYPTONERRCODE0047");
                            break;
                        }
                    }
                }
                if (!validation)
                {
                    throw new Exception(errorMessage);
                }

            }
            catch (Exception e)
            {
                Common.KryptonException.writeexception(e);
                throw new Common.KryptonException(e.Message);
            }
        }

        /// <summary>
        /// Store the Streamreader comtents to runtimedic dictionary.
        /// </summary>
        /// <param name="reader">StreamReader object </param>
        public static void StoreReaderContent(StreamReader reader)
        {
            try
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    int seperatorPos = line.IndexOf(Common.Property.KeyValueDistinctionKeyword);
                    if (seperatorPos > 0)
                    {
                        string[] param = line.Split(Common.Property.KeyValueDistinctionKeyword);

                        string val = string.Empty;
                        if (param.Length > 2)
                        {
                            for (int paramIndex = 1; paramIndex < param.Length; paramIndex++)
                            {
                                if (string.IsNullOrWhiteSpace(val)) val = param[paramIndex];
                                else val = val + Common.Property.KeyValueDistinctionKeyword + param[paramIndex];
                            }
                        }
                        else
                        {
                            val = param[1];
                        }

                        val = val.Replace("\"", "");
                        val = val.Replace(@"./", @".\");

                        try
                        {

                            if (Common.Property.parameterdic.ContainsKey(param[0].ToLower().Trim()))
                            {
                                Common.Property.parameterdic[param[0].ToLower().Trim()] = val.Trim();
                                Common.Property.runtimedic[param[0].ToLower().Trim()] = val.Trim();
                            }
                            else
                            {
                                Common.Property.parameterdic.Add(param[0].ToLower().Trim(), val.Trim());
                                Common.Property.runtimedic.Add(param[0].ToLower().Trim(), val.Trim());
                            }
                        }
                        catch
                        {
                            //No throw
                        }


                    }
                }

                if (Utility.GetParameter("projectpath").ToLower().Equals("true"))  // set projectPath variable for using it as a variable anywhere
                {
                    SetParameter("projectpath", Property.IniPath);
                    SetVariable("projectpath", Property.IniPath);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Initialize Browser for execution.
        /// </summary>
        /// <param name="browserrun">Browser on which execution will run.</param>
        public static void InitializeBrowser(string browserrun)
        {
            BrowserManager.getBrowser(browserrun);
        }

        /// <summary>
        ///Email validation method to validate the format of email files.
        /// </summary>
        /// <returns>true on success</returns>
        private static bool ValidateEmail()
        {
            long num;
            if (GetParameter("EmailNotification").Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                if (GetParameter("EmailNotificationFrom") == string.Empty || !GetParameter("EmailNotificationFrom").Contains("@"))
                {
                    Console.WriteLine("Warning: \"EmailNotificationFrom\" parameter is invalid! Email notification process aborted");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPServer") == string.Empty || !GetParameter("EmailSMTPServer").Contains("."))
                {
                    Console.WriteLine("Warning: \"EmailSMTPServer\" parameter is invalid! Email notification process aborted");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPPort") == string.Empty || !long.TryParse(GetParameter("EmailSMTPPort"), out num))
                {
                    Console.WriteLine("Warning: \"EmailSMTPPort\" parameter is invalid! Email notification process aborted");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPUsername") == string.Empty || !GetParameter("EmailSMTPUsername").Contains("@"))
                {
                    Console.WriteLine("Warning: \"EmailSMTPUsername\" parameter is invalid! Email notification process aborted");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPPassword") == string.Empty)
                {
                    Console.WriteLine("Warning: \"EmailSMTPPassword\" parameter is invalid! Email notification process aborted");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
                // so that templates folder will accept both absolute and relative paths.
                //-----------------------------Start Modifiaction---------------------------------
                string startTemplatePath = Property.EmailStartTemplate;
                if (!File.Exists(startTemplatePath))
                {
                    startTemplatePath = Common.Property.ApplicationPath + GetParameter("EmailStartTemplate");
                }
                if (!File.Exists(startTemplatePath))
                {
                    Console.WriteLine("Warning: Email StartTemplate file is missing! Email notification process aborted");
                    Console.WriteLine("------------------------------------------------");
                    return false;
                }

                string endTemplatePath = Property.EmailEndTemplate;
                if (!File.Exists(endTemplatePath))
                {
                    endTemplatePath = Common.Property.ApplicationPath + GetParameter("EmailEndTemplate");
                }
                if (!File.Exists(endTemplatePath))
                {
                    Console.WriteLine("Warning: Email EndTemplate file is missing! Email notification process aborted");
                    Console.WriteLine("------------------------------------------------");
                    return false;
                }
                string startTemplate = File.ReadAllText(startTemplatePath);
                if (!startTemplate.ToLower().Contains("subject::") || !startTemplate.ToLower().Contains("body::"))
                {
                    Console.WriteLine("Warning: Email StartTemplate format is invalid! Email notification process aborted");
                    Console.WriteLine("------------------------------------------------");
                    return false;
                }
                string endTemplate = File.ReadAllText(endTemplatePath);
                if (!endTemplate.ToLower().Contains("subject::") || !endTemplate.ToLower().Contains("body::"))
                {
                    Console.WriteLine("Warning: Email EndTemplate format is invalid! Email notification process aborted");
                    Console.WriteLine("------------------------------------------------");
                    return false;
                }
                //-----------------------------end Modifiaction---------------------------------
                string emailNotif = File.ReadAllText(Path.Combine(Common.Property.IniPath, "EmailNotification.ini"));
                if (!emailNotif.Trim().ToLower().Contains("recipient"))
                {
                    Console.WriteLine("Warning: EmailNotification.ini format is invalid! Email notification process aborted");
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        ///Email Notification method for different stages
        /// </summary>
        /// <param name="stage">stage/state of execution</param>
        public static void EmailNotificationOld(string stage, bool zipAttch, string reportFileName = null)
        {
            if (ValidateEmail() == false) return;

            Console.WriteLine(ConsoleMessages.MSG_DASHED);
            Console.WriteLine("Sending Email....");
            Console.WriteLine(ConsoleMessages.MSG_DASHED);

            try
            {
                string templateFileName = string.Empty;
                string[] emailStr = null;
                switch (stage.ToLower())
                {

                    case "start":
                        emailStr = EmailTemplate(Property.EmailStartTemplate);
                        break;
                    case "end":
                        emailStr = EmailTemplate(Property.EmailEndTemplate);
                        break;
                    default:
                        break;
                }

                MailMessage message = new MailMessage();

                MailAddress mailFrom = new MailAddress(GetParameter("EmailNotificationFrom"), "Krypton Automation");

                message.From = mailFrom;

                string mailSubject = string.Empty;
                mailSubject = emailStr[1];
                if (string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false)
                    mailSubject = mailSubject.Replace("[Test Suite]", GetParameter("TestSuite"));
                else
                    mailSubject = mailSubject.Replace("[Test Suite]", string.Empty);
                if (Property.FinalExecutionStatus.ToLower().Equals("pass"))
                {
                    Property.FinalExecutionStatus = "Passed";
                }
                if (Property.FinalExecutionStatus.ToLower().Equals("fail"))
                {
                    Property.FinalExecutionStatus = "Failed";
                }
                mailSubject = mailSubject.Replace("[Result]", Property.FinalExecutionStatus);
                mailSubject = mailSubject.Replace("[StartTime]", Property.ExecutionStartDateTime);
                mailSubject = mailSubject.Replace("[EndTime]", Property.ExecutionEndDateTime);

                mailSubject = mailSubject.Replace("[TotalStepExecuted]", Property.TotalStepExecuted.ToString());
                mailSubject = mailSubject.Replace("[TotalStepPass]", Property.TotalStepPass.ToString());
                mailSubject = mailSubject.Replace("[TotalStepFail]", Property.TotalStepFail.ToString());
                mailSubject = mailSubject.Replace("[TotalStepWarning]", Property.TotalStepWarning.ToString());

                mailSubject = mailSubject.Replace("[TotalCaseExecuted]", Property.TotalCaseExecuted.ToString());
                emailStr[3] = emailStr[3].Replace("[TotalCaseExecuted]", Property.TotalCaseExecuted.ToString());
                if (Property.FinalExecutionStatus.ToLower().Equals("passed"))
                    mailSubject = mailSubject.Replace("[TotalCaseFail]/", string.Empty);
                else
                    mailSubject = mailSubject.Replace("[TotalCaseFail]", Property.TotalCaseFail.ToString());

                mailSubject = mailSubject.Replace("[TotalCasePass]", Property.TotalCasePass.ToString());
                mailSubject = mailSubject.Replace("[TotalCaseWarning]", Property.TotalCaseWarning.ToString());
                SmtpClient smtpClient = new SmtpClient(GetParameter("EmailSMTPServer"), int.Parse(GetParameter("EmailSMTPPort")));
                NetworkCredential credential = new NetworkCredential(GetParameter("EmailSMTPUsername"), GetParameter("EmailSMTPPassword"));
                smtpClient.Credentials = credential;
                smtpClient.EnableSsl = true;
                /// send mail to comma seperated email recipients -------
                if (File.Exists(Path.Combine(Common.Property.IniPath, Property.EmailNotificationFile)))
                {
                    try
                    {
                        string allrecipients = File.ReadAllText(Path.Combine(Common.Property.IniPath, Property.EmailNotificationFile));
                        allrecipients = allrecipients.Substring(allrecipients.IndexOf(':') + 1);
                        string[] recipient = allrecipients.Split(',');
                        {
                            foreach (string recipientAdd in recipient)
                            {
                                if ((!string.IsNullOrWhiteSpace(recipientAdd)) && (!string.IsNullOrEmpty(recipientAdd)))
                                    message.To.Add(new MailAddress(recipientAdd.Trim(), null));
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        throw new Exception("Invalid Email format in EmailNotification.ini");
                    }

                }
                try
                {

                    if (string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false)
                        emailStr[3] = emailStr[3].Replace("[Test Suite]", GetParameter("TestSuite"));
                    else
                        emailStr[3] = emailStr[3].Replace("[Test Suite]", string.Empty);
                    emailStr[3] = emailStr[3].Replace("[Result]", Property.FinalExecutionStatus);
                    emailStr[3] = emailStr[3].Replace("[StartTime]", Property.ExecutionStartDateTime);
                    if (stage == "end")
                    {
                        emailStr[3] = emailStr[3].Replace("[EndTime]", Property.ExecutionEndDateTime);

                        DateTime executionStartTime = DateTime.ParseExact(Property.ExecutionStartDateTime, Common.Utility.GetParameter("DateTimeFormat"), null);
                        DateTime executionEndTime = DateTime.ParseExact(Property.ExecutionEndDateTime, Common.Utility.GetParameter("DateTimeFormat"), null);
                        TimeSpan time = executionEndTime - executionStartTime;

                        emailStr[3] = emailStr[3].Replace("[ExecutionTime]", time.ToString());
                    }
                    if (GetParameter("RunRemoteExecution") == "true")
                        emailStr[3] = emailStr[3].Replace("[RemoteUrl]", GetParameter("RunOnRemoteBrowserUrl"));
                    else
                        emailStr[3] = emailStr[3].Replace("[RemoteUrl]", "localhost");

                    emailStr[3] = emailStr[3].Replace("[Environment]", GetParameter("Environment"));

                    emailStr[3] = emailStr[3].Replace("[TotalStepExecuted]", Property.TotalStepExecuted.ToString());
                    emailStr[3] = emailStr[3].Replace("[TotalStepPass]", Property.TotalStepPass.ToString());
                    emailStr[3] = emailStr[3].Replace("[TotalStepFail]", Property.TotalStepFail.ToString());
                    emailStr[3] = emailStr[3].Replace("[TotalStepWarning]", Property.TotalStepWarning.ToString());

                    emailStr[3] = emailStr[3].Replace("[TotalCaseExecuted]", Property.TotalCaseExecuted.ToString());
                    if (Property.FinalExecutionStatus.ToLower().Equals("passed"))
                        emailStr[3] = emailStr[3].Replace("[TotalCaseFail]/", string.Empty);
                    else
                        emailStr[3] = emailStr[3].Replace("[TotalCaseFail]", Property.TotalCaseFail.ToString());

                    emailStr[3] = emailStr[3].Replace("[TotalCasePass]", Property.TotalCasePass.ToString());
                    emailStr[3] = emailStr[3].Replace("[TotalCaseWarning]", Property.TotalCaseWarning.ToString());

                    //Results Destination Path variable
                    emailStr[3] = emailStr[3].Replace("[LogDestinationFolder]", Property.ResultsDestinationPath.Substring(0, Property.ResultsDestinationPath.LastIndexOf('\\')));


                    Attachment data = null;

                    if (stage == "end")
                    {
                        // Create  the file attachment for this e-mail message.
                        string file = Common.Property.HtmlFileLocation + "/" + Property.ReportZipFileName;

                        int doStep = 0;
                        do
                        {
                            if (File.Exists(file) && zipAttch)
                            {
                                FileInfo fileInfo = new FileInfo(file);
                                long fileSize = fileInfo.Length;
                                if (fileSize > 1024 * 1024 * 4)
                                {
                                    emailStr[3] = emailStr[3] +
                                                  "\n---------------\n**NOTE: Not able to send attachement, size of the attachment is too big.**";

                                    string htmlFile = Common.Property.HtmlFileLocation + "/HtmlReport.html";
                                    data = new Attachment(htmlFile, MediaTypeNames.Application.Octet);
                                    // Add time stamp information for the file.
                                    ContentDisposition htmlDisposition = data.ContentDisposition;
                                    htmlDisposition.CreationDate = System.IO.File.GetCreationTime(htmlFile);
                                    htmlDisposition.ModificationDate = System.IO.File.GetLastWriteTime(htmlFile);
                                    htmlDisposition.ReadDate = System.IO.File.GetLastAccessTime(htmlFile);
                                    // Add the html file attachment to this e-mail message.
                                    message.Attachments.Add(data);

                                    break;
                                }

                                data = new Attachment(file, MediaTypeNames.Application.Octet);
                                // Add time stamp information for the file.
                                ContentDisposition disposition = data.ContentDisposition;
                                disposition.CreationDate = System.IO.File.GetCreationTime(file);
                                disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                                disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                                // Add the zip file attachment to this e-mail message.
                                message.Attachments.Add(data);

                                break;
                            }
                            else if (zipAttch == false)
                            {
                                string htmlFile = string.Empty;
                                if (string.IsNullOrWhiteSpace(reportFileName))
                                {
                                    htmlFile = Common.Property.HtmlFileLocation + "/HtmlReportMail.html";

                                }
                                else
                                {
                                    htmlFile = reportFileName;
                                }

                                data = new Attachment(htmlFile, MediaTypeNames.Application.Octet);
                                // Add time stamp information for the file.
                                ContentDisposition htmlDisposition = data.ContentDisposition;
                                htmlDisposition.CreationDate = System.IO.File.GetCreationTime(htmlFile);
                                htmlDisposition.ModificationDate = System.IO.File.GetLastWriteTime(htmlFile);
                                htmlDisposition.ReadDate = System.IO.File.GetLastAccessTime(htmlFile);
                                // Add the html file attachment to this e-mail message.
                                message.Attachments.Add(data);

                                break;
                            }
                            else
                            {
                                doStep++;
                                if (doStep < 5)
                                    Thread.Sleep(5000); //wait for 5 sec.
                            }
                        } while (doStep < 5);

                    }

                    message.Subject = mailSubject;
                    if (!string.IsNullOrEmpty(Common.Property.ReportSummaryBody))
                    {
                        message.Body = Common.Property.ReportSummaryBody;
                        message.IsBodyHtml = true;
                    }
                    else
                    {
                        message.Body = emailStr[3];
                        message.IsBodyHtml = false;
                    }
                    smtpClient.Send(message);

                    message.Dispose();
                    smtpClient.Dispose();

                }
                catch
                {
                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to send email notificiation. Error: " + ex.Message + "\n" +
                                  "Please make sure you have smtp service installed and running");

            }
        }

        /// <summary>
        ///Updated Method,Send Mail with CDO method
        /// </summary>
        /// <param name="stage"></param>
        /// <param name="zipAttch"></param>
        /// <param name="reportFileName"></param>
        public static void EmailNotification(string stage, bool zipAttch, string reportFileName = null)
        {

            if (ValidateEmail() == false) return;

            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("Sending Email ....!!!");
            Console.WriteLine("------------------------------------------------");

            try
            {
                string templateFileName = string.Empty;
                string[] emailStr = null;
                switch (stage.ToLower())
                {

                    case "start":
                        emailStr = EmailTemplate(Property.EmailStartTemplate);
                        break;
                    case "end":
                        emailStr = EmailTemplate(Property.EmailEndTemplate);
                        break;
                    default:
                        break;
                }


                #region Email Configuration settings

                CDO.Message message = new CDO.Message();
                CDO.IConfiguration configuration = message.Configuration;
                ADODB.Fields fields = configuration.Fields;
                ADODB.Field field = fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"];
                field.Value = GetParameter("EmailSMTPServer").Trim(); //"smtp.gmail.com"; 

                field = fields["http://schemas.microsoft.com/cdo/configuration/smtpserverport"];
                field.Value = int.Parse(GetParameter("EmailSMTPPort").Trim());// should be 465

                field = fields["http://schemas.microsoft.com/cdo/configuration/sendusing"];
                field.Value = CDO.CdoSendUsing.cdoSendUsingPort;

                field = fields["http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"];
                field.Value = CDO.CdoProtocolsAuthentication.cdoBasic;

                field = fields["http://schemas.microsoft.com/cdo/configuration/sendusername"];
                field.Value = GetParameter("EmailSMTPUsername").Trim(); //"krypton.thinksys@gmail.com";

                field = fields["http://schemas.microsoft.com/cdo/configuration/sendpassword"];
                field.Value = GetParameter("EmailSMTPPassword").Trim(); //"Thinksys@123";  

                field = fields["http://schemas.microsoft.com/cdo/configuration/smtpusessl"];
                field.Value = "true";
                fields.Update();

                #endregion



                #region   Add attachments

                if (stage == "end")
                {
                    // Create  the file attachment for this e-mail message.
                    string file = Common.Property.HtmlFileLocation + "/" + Property.ReportZipFileName;
                    if (File.Exists(file) && zipAttch)
                    {
                        message.AddAttachment(file);

                    }
                    else if (!zipAttch)
                    {
                        string htmlFile = string.Empty;
                        if (string.IsNullOrWhiteSpace(reportFileName))
                        {
                            htmlFile = Common.Property.HtmlFileLocation + "/HtmlReportMail.html";

                        }
                        else
                        {
                            htmlFile = reportFileName;
                        }

                        message.AddAttachment(htmlFile);
                    }

                }

                #endregion

                if (!string.IsNullOrEmpty(Common.Property.ReportSummaryBody))
                {
                    message.HTMLBody = Common.Property.ReportSummaryBody;

                }
                else
                {
                    string mailTextBody = GetEmailTextBody(emailStr[3], stage);
                    message.TextBody = mailTextBody;

                }
                message.From = "KryptonAutomation";
                message.Sender = GetParameter("EmailSMTPUsername").Trim();
                message.To = GetRecipientList();
                message.Subject = GetEmailSubjectMessage(emailStr[1]); ;
                message.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to send email notificiation. Error: " + ex.Message + "\n" +
                                  "Please make sure you have smtp service installed and running");

            }

        }

        /// <summary>
        ///Returns Email Text Body if no Html body is available
        /// </summary>
        /// <param name="emailTextBody"></param>
        /// <param name="stage"></param>
        /// <returns></returns>
        private static string GetEmailTextBody(string emailTextBody, string stage)
        {
            string EmailTextBody = string.Empty;
            try
            {
                if (string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false)
                    emailTextBody = emailTextBody.Replace("[Test Suite]", GetParameter("TestSuite"));
                else
                    emailTextBody = emailTextBody.Replace("[Test Suite]", string.Empty);
                emailTextBody = emailTextBody.Replace("[Result]", Property.FinalExecutionStatus);
                emailTextBody = emailTextBody.Replace("[StartTime]", Property.ExecutionStartDateTime);
                if (stage == "end")
                {
                    emailTextBody = emailTextBody.Replace("[EndTime]", Property.ExecutionEndDateTime);

                    DateTime executionStartTime = DateTime.ParseExact(Property.ExecutionStartDateTime, Common.Utility.GetParameter("DateTimeFormat"), null);
                    DateTime executionEndTime = DateTime.ParseExact(Property.ExecutionEndDateTime, Common.Utility.GetParameter("DateTimeFormat"), null);
                    TimeSpan time = executionEndTime - executionStartTime;

                    emailTextBody = emailTextBody.Replace("[ExecutionTime]", time.ToString());
                }
                if (GetParameter("RunRemoteExecution") == "true")
                    emailTextBody = emailTextBody.Replace("[RemoteUrl]", GetParameter("RunOnRemoteBrowserUrl"));
                else
                    emailTextBody = emailTextBody.Replace("[RemoteUrl]", "localhost");

                emailTextBody = emailTextBody.Replace("[Environment]", GetParameter("Environment"));

                emailTextBody = emailTextBody.Replace("[TotalStepExecuted]", Property.TotalStepExecuted.ToString());
                emailTextBody = emailTextBody.Replace("[TotalStepPass]", Property.TotalStepPass.ToString());
                emailTextBody = emailTextBody.Replace("[TotalStepFail]", Property.TotalStepFail.ToString());
                emailTextBody = emailTextBody.Replace("[TotalStepWarning]", Property.TotalStepWarning.ToString());

                emailTextBody = emailTextBody.Replace("[TotalCaseExecuted]", Property.TotalCaseExecuted.ToString());
                if (Property.FinalExecutionStatus.ToLower().Equals("passed"))
                    emailTextBody = emailTextBody.Replace("[TotalCaseFail]/", string.Empty);
                else
                    emailTextBody = emailTextBody.Replace("[TotalCaseFail]", Property.TotalCaseFail.ToString());

                emailTextBody = emailTextBody.Replace("[TotalCasePass]", Property.TotalCasePass.ToString());
                emailTextBody = emailTextBody.Replace("[TotalCaseWarning]", Property.TotalCaseWarning.ToString());

                //Results Destination Path variable
                EmailTextBody = emailTextBody = emailTextBody.Replace("[LogDestinationFolder]", Property.ResultsDestinationPath.Substring(0, Property.ResultsDestinationPath.LastIndexOf('\\')));
                return EmailTextBody;
            }
            catch
            {

                return string.Empty;
            }

        }
        /// <summary>
        ///Return Email Subject
        /// </summary>
        /// <param name="EmailMessage"></param>
        /// <returns></returns>
        private static string GetEmailSubjectMessage(string EmailMessage)
        {
            string mailSubject = string.Empty;
            try
            {
                mailSubject = EmailMessage;
                if (string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false)
                    mailSubject = mailSubject.Replace("[Test Suite]", GetParameter("TestSuite"));
                else
                    mailSubject = mailSubject.Replace("[Test Suite]", "");
                if (Property.FinalExecutionStatus.ToLower().Equals("pass"))
                {
                    Property.FinalExecutionStatus = "Passed";
                }
                if (Property.FinalExecutionStatus.ToLower().Equals("fail"))
                {
                    Property.FinalExecutionStatus = "Failed";
                }
                mailSubject = mailSubject.Replace("[Result]", Property.FinalExecutionStatus);
                mailSubject = mailSubject.Replace("[StartTime]", Property.ExecutionStartDateTime);
                mailSubject = mailSubject.Replace("[EndTime]", Property.ExecutionEndDateTime);

                mailSubject = mailSubject.Replace("[TotalStepExecuted]", Property.TotalStepExecuted.ToString());
                mailSubject = mailSubject.Replace("[TotalStepPass]", Property.TotalStepPass.ToString());
                mailSubject = mailSubject.Replace("[TotalStepFail]", Property.TotalStepFail.ToString());
                mailSubject = mailSubject.Replace("[TotalStepWarning]", Property.TotalStepWarning.ToString());
                mailSubject = mailSubject.Replace("[TotalCaseExecuted]", Property.TotalCaseExecuted.ToString());
                if (Property.FinalExecutionStatus.ToLower().Equals("passed"))
                    mailSubject = mailSubject.Replace("[TotalCaseFail]/", string.Empty);
                else
                    mailSubject = mailSubject.Replace("[TotalCaseFail]", Property.TotalCaseFail.ToString());
                mailSubject = mailSubject.Replace("[TotalCasePass]", Property.TotalCasePass.ToString());
                mailSubject = mailSubject.Replace("[TotalCaseWarning]", Property.TotalCaseWarning.ToString());
                return mailSubject;
            }
            catch
            {
                return string.Empty;
            }

        }

        /// <summary>
        ///Prepare Email recipient List.
        /// 
        /// </summary>
        /// <returns> Comma Seperated email recepients list</returns>
        private static string GetRecipientList()
        {
            string RecipientList = string.Empty;
            if (File.Exists(Path.Combine(Common.Property.IniPath, Property.EmailNotificationFile)))
            {
                try
                {
                    string allrecipients = File.ReadAllText(Path.Combine(Common.Property.IniPath, Property.EmailNotificationFile));
                    allrecipients = allrecipients.Substring(allrecipients.IndexOf(':') + 1);
                    string[] recipient = allrecipients.Split(',');
                    {
                        foreach (string recipientAdd in recipient)
                        {
                            if ((!string.IsNullOrWhiteSpace(recipientAdd)) && (!string.IsNullOrEmpty(recipientAdd)))
                                RecipientList = RecipientList + "," + recipientAdd.Trim();

                        }
                    }
                }
                catch (Exception e)
                {
                    throw new Exception("Invalid Email format in EmailNotification.ini");
                }
            }
            return RecipientList;
        }

        /// <summary>
        ///Email Notification template capturing
        /// </summary>
        /// <param name="fileName">File name of the template</param>
        /// <returns>string array value of the subject and body</returns>
        private static string[] EmailTemplate(string fileName)
        {
            FileInfo emailFileInfo = new FileInfo(fileName);
            string[] emailTemplate = Regex.Split(File.ReadAllText(emailFileInfo.FullName), "::");

            return emailTemplate;
        }

        /// <summary>
        /// Get Option type value based upon passing parameter like testmode, testdatasheet etc.
        /// 
        /// </summary>
        /// <param name="optionValue">{testmode=A}{testdatasheet="login"} etc.</param>
        /// <param name="optionType">testmode, testdatasheet</param>
        /// <returns>string test mode</returns>
        public static string GetTestMode(string optionValue, string optionType)
        {
            string testModeStr = string.Empty;

            try
            {
                for (int v = 0; ; v++)
                {
                    if (optionValue.Contains("{"))
                    {
                        int stindex = optionValue.IndexOf("{");
                        optionValue = optionValue.Remove(stindex, 1);

                        int endindex = optionValue.IndexOf("}");
                        if (endindex < 0) break;

                        string KeyVariable = optionValue.Substring(stindex, (endindex - stindex));
                        // Modified to check the existence of 'script' or 'testmode' keywords before '=' only.
                        if (KeyVariable.IndexOf(optionType, StringComparison.OrdinalIgnoreCase) >= 0 && KeyVariable.Contains("="))
                            if (KeyVariable.IndexOf(optionType, StringComparison.OrdinalIgnoreCase) < KeyVariable.IndexOf("="))
                                testModeStr = KeyVariable.Split('=')[0].Trim().Replace("\"", "") + "," + KeyVariable.Split('=')[1].Trim().Replace("\"", "");
                        stindex = optionValue.IndexOf("}");
                        optionValue = optionValue.Remove(stindex, 1);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception testMode)
            {
                Console.WriteLine("Exception while determining test mode: " + testMode.Message);
            }

            return testModeStr;
        }

        // get OR row in form of Dictionary
        public static Dictionary<string, string> GetTestOrData(string parent, string testObj, DataSet ORTestData)
        {
            Dictionary<string, string> orDataRow = new Dictionary<string, string>();
            if (orDataRow.ContainsKey("logical_name"))
            {
                orDataRow["logical_name"] = string.Empty;
                orDataRow[KryptonConstants.OBJ_TYPE] = string.Empty;
                orDataRow[KryptonConstants.HOW] = string.Empty;
                orDataRow[KryptonConstants.WHAT] = string.Empty;
                orDataRow[KryptonConstants.MAPPING] = string.Empty;
            }
            try
            {

                foreach (DataRow drData in ORTestData.Tables[0].Rows)
                {

                    if (drData["parent"].ToString().Equals(parent, StringComparison.OrdinalIgnoreCase)
                        && drData[KryptonConstants.TEST_OBJECT].ToString().Equals(testObj, StringComparison.OrdinalIgnoreCase)
                        )
                    {
                        if (orDataRow.ContainsKey("logical_name") == false)
                        {
                            orDataRow.Add("parent", Common.Utility.ReplaceVariablesInString(drData["parent"].ToString().Trim()));
                            orDataRow.Add("logical_name", Common.Utility.ReplaceVariablesInString(drData["logical_name"].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.OBJ_TYPE, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.HOW, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.WHAT, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString().Trim()));
                            orDataRow.Add(KryptonConstants.MAPPING, Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString().Trim()));
                        }
                        else
                        {
                            orDataRow.Add("parent", Common.Utility.ReplaceVariablesInString(drData["parent"].ToString().Trim()));
                            orDataRow["logical_name"] = Common.Utility.ReplaceVariablesInString(drData["logical_name"].ToString());
                            orDataRow[KryptonConstants.OBJ_TYPE] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString());
                            orDataRow[KryptonConstants.HOW] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString());
                            orDataRow[KryptonConstants.WHAT] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString());
                            orDataRow[KryptonConstants.MAPPING] = Common.Utility.ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString());
                        }
                    }

                }
            }
            catch (Exception e)
            {
                throw e;
            }


            return orDataRow;
        }

        /// <summary>
        ///Method to copy all the contents of source directory to destination directory recursively
        /// </summary>
        /// <param name="source">source directory</param>
        /// <param name="target">destination directory</param>
        /// <returns>true on successful copy</returns>
        public static bool CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            try
            {
                //check whether target directory exists and create it if not present.
                if (Directory.Exists(target.FullName) == false)
                {
                    Directory.CreateDirectory(target.FullName);
                }

                //copy all the files into source directory to target directory.
                foreach (FileInfo fi in source.GetFiles())
                {
                    try
                    {
                        fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);
                    }
                    catch
                    { /*leave the locked files as they are not required for profile*/ }
                }

                //find all the directories into source directory and make a recursive call to copy its file to destination folder.
                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {
                    DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                    CopyAll(diSourceSubDir, nextTargetSubDir);
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }

            return true;
        }

        /// <summary>
        ///Method to make a backup of firefox profile after execution of test case completed.
        ///This is used to keep the cookies stored after completion of test case.
        /// </summary>
        public static void firefoxBackup()
        {
            if (GetParameter("FirefoxProfilePath").Length != 0 && GetParameter("Browser").Equals("firefox", StringComparison.OrdinalIgnoreCase))
            {
                // Create backup of firefox profile
                DirectoryInfo ffProfile = new DirectoryInfo(GetParameter("FirefoxProfilePath"));
                DirectoryInfo ffProfileBackup = new DirectoryInfo(GetParameter("FirefoxProfilePath") + "_bkp");
                if (ffProfile.Exists)
                    CopyAll(ffProfile, ffProfileBackup);
            }
        }

        /// <summary>
        ///Method to clean up the temporary profile and addons directories after completion of test case.
        /// </summary>
        public static void firefoxCleanup()
        {
            if (GetParameter("FirefoxProfilePath").Length != 0 && GetParameter("Browser").Equals("firefox", StringComparison.OrdinalIgnoreCase))
            {
                //we are reverting back to b3
                if (Directory.Exists(Common.Utility.GetParameter("tmpFFProfileDir")))
                    Directory.Delete(Common.Utility.GetParameter("tmpFFProfileDir"), true);
            }

            if (GetParameter("AddonsPath").Length != 0 && GetParameter("Browser").Equals("firefox", StringComparison.OrdinalIgnoreCase))
            {
                DirectoryInfo addonDir = new DirectoryInfo("AddonsTemp");
                if (addonDir.Exists)
                    addonDir.Delete(true);
            }
        }


        /// <summary>
        /// This method will return Clean up test case id with options
        /// </summary>
        /// <param name="optionValue"></param>
        /// <param name="optionType"></param>
        /// <returns></returns>
        public static string GetCleanupTestCase(string optionValue, string optionType)
        {
            string testModeStr = string.Empty;


            for (int v = 0; ; v++)
            {
                if (optionValue.Contains("{"))
                {
                    int stindex = optionValue.IndexOf("{");
                    if (stindex > -1)
                    {
                        optionValue = optionValue.Remove(stindex, 1);
                        int endindex = optionValue.IndexOf("}");
                        if (endindex > -1)
                        {
                            string KeyVariable = optionValue.Substring(stindex, (endindex - stindex));
                            if (KeyVariable.IndexOf(optionType, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                testModeStr = KeyVariable;
                            }
                            stindex = optionValue.IndexOf("}");
                            optionValue = optionValue.Remove(stindex, 1);
                        }
                    }
                }
                else
                {
                    break;
                }
            }
            return testModeStr;
        }

        /// <summary>
        ///this method is to clear run time dictionary and fill parameter dictionary in run time dictionary
        /// </summary>
        public static void ClearRunTimeDic()
        {
            Property.runtimedic.Clear();
            Property.runtimedic = new Dictionary<string, string>(Property.parameterdic);

            return;
        }

        /// <summary>
        ///Creating temporary file
        /// </summary>
        /// <param name="extn"></param>
        /// <returns></returns>
        public static string GetTemporaryFile(string extn)
        {
            string response = string.Empty;

            if (!extn.StartsWith("."))
                extn = "." + extn;

            response = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + extn;

            return response;
        }


        /// <summary>
        ///This method will fetch test data from Test Case Excel Sheet into a dataset
        /// and will return the dataset.
        /// It will get the test case sheet location from GetTestFilesLocation method.
        /// </summary>
        /// <returns>Dataset containing Test Data.</returns>
        public static void GetCommonMessageData()
        {
            string oriFilePath = Common.Property.ApplicationPath + Property.GlobalCommentsFile;
            string tmpFileName = string.Empty;
            tmpFileName = Utility.GetTemporaryFile(".txt");
            if (File.Exists(oriFilePath))
            {
                File.Copy(oriFilePath, tmpFileName);

                if (File.Exists(tmpFileName))
                {
                    FileStream fileStream = new FileStream(tmpFileName, FileMode.Open, FileAccess.Read, FileShare.None);
                    StreamReader streamReader = new StreamReader(fileStream);
                    string readLine;
                    while ((readLine = streamReader.ReadLine()) != null)
                    {
                        int splitInd = readLine.IndexOf('|');
                        string messageId = readLine.Substring(0, splitInd);
                        string messageStr = readLine.Substring(splitInd + 1, readLine.Length - splitInd - 1);
                        Utility.SetCommonMsgVariable(messageId, messageStr);
                    }
                    Property.listOfFilesInTempFolder.Add(tmpFileName);
                }
                else
                {
                    throw new Common.KryptonException("Error", "File missing:" + tmpFileName);
                }

            }
            else
            {
                throw new Common.KryptonException("Error", "File missing:" + oriFilePath);
            }


        }

        /// <summary>
        ///Stores a variable in common message dictionary. Variable name in dictionary are always lower case.
        ///This will also overwrite any previously written variable value.
        /// </summary>
        /// <param name="varName"></param>
        /// <param name="varValue"></param>
        public static void SetCommonMsgVariable(string varName, string varValue)
        {
            varName = varName.ToLower();
            if (Common.Property.commonMsgdic.ContainsKey(varName))
                Common.Property.commonMsgdic[varName] = varValue;
            else
                Common.Property.commonMsgdic.Add(varName, varValue);
        }


        /// <summary>
        ///Retrives value of variable from common message dictionary. Return ZERO length string if variable
        /// could not be found in dictionary.
        /// </summary>
        /// <param name="varName"></param>
        /// <returns></returns>
        public static string GetCommonMsgVariable(string varName)
        {
            varName = varName.ToLower();
            string value = string.Empty;
            try
            {
                value = Common.Property.commonMsgdic[varName];
            }
            catch (Exception)
            {
                return varName;//return variable name itself if given varaible has no value.
            }
            return value;
        }

        /// <summary>
        ///Replaces special characters in a string. Following special characters are replaced with actual characters:
        ///{PIPE}, {SPACE}, {NONE}
        /// </summary>
        /// <param name="inString">Input string, in which special character needs to be replaced.</param>
        /// <returns>String with special characters replaced</returns>
        public static string ReplaceSpecialCharactersInString(string inString)
        {
            if (inString.Contains("{PIPE}"))
                inString = inString.Replace("{PIPE}", "|");

            if (inString.Contains("{SPACE}"))
                inString = inString.Replace("{SPACE}", " ");

            if (inString.Contains("{TAB}"))
                inString = inString.Replace("{TAB}", "\t");

            if (inString.Contains("{NONE}"))
                inString = inString.Replace("{NONE}", "");

            return inString;
        }

        private static string[] getRegistryKey(string softwareKey)
        {
            RegistryKey rk = Registry.LocalMachine.OpenSubKey(softwareKey);
            return rk.GetSubKeyNames();
        }

        /// <summary>
        /// Get Sub Key String For firefox browser.
        /// </summary>
        /// <returns>string : subKey String.</returns>
        public static string getFFVersionString()
        {
            string softwareKey = Property.Browser_KeyString;
            string versionString = string.Empty;
            try
            {
                string[] subKeys = getRegistryKey(softwareKey);
                foreach (string key in subKeys)
                {
                    if (key.Contains("Mozilla Firefox"))
                        versionString = key;
                }
                //For Win 7 machines.
                if (versionString.Equals(String.Empty))
                {
                    string[] keys = getRegistryKey(Property.Browser_KeyString_64bit);
                    foreach (string key in keys)
                    {
                        if (key.Contains("Mozilla Firefox"))
                            versionString = key;
                    }
                }
                return versionString;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
        /// <summary>
        /// Register any dll on Application machine.
        /// </summary>
        /// <param name="dllPath">string : Path to Dll.</param>
        public static void registerDll(string dllPath)
        {
            //registering dll silentltly.
            string fileinfo = "/s" + " " + "\"" + dllPath + "\"";
            try
            {
                Process reg = new Process();
                reg.StartInfo.FileName = "regsvr32.exe";
                reg.StartInfo.Arguments = fileinfo;
                reg.StartInfo.UseShellExecute = false;
                reg.StartInfo.CreateNoWindow = true;
                reg.StartInfo.RedirectStandardOutput = true;
                reg.Start();
                reg.WaitForExit();
                reg.Close();
            }
            catch (Exception)
            {

            }
        }

        /// <summary>
        /// Method to retrieve all numbers from given string.
        /// </summary>
        /// <param name="text">string </param>
        /// <returns>number</returns>
        public static int GetNumberFromText(string text)
        {
            string number = string.Empty;
            foreach (char c in text)
            {
                int a = System.Convert.ToInt32(c);
                if (a >= 48 && a <= 57)
                {
                    number = number + c.ToString();

                }
            }

            return Int32.Parse(number);
        }

        #region random number generation
        /// <summary>
        ///These steps need to be written to have syncronize random number generation
        /// so that number will be trully random
        /// </summary>

        private static readonly Random random = new Random();
        private static readonly object syncLock = new object();
        public static int RandomNumber(int min, int max)
        {
            lock (syncLock)
            {   // synchronize
                return random.Next(min, max);
            }
        }

        public static string RandomString(int strLength)
        {
            string randomStr = string.Empty;
            for (int chrCnt = 0; chrCnt < strLength; chrCnt++)
            {
                randomStr = randomStr + Property.ListOfUniqueCharacters[RandomNumber(0, Property.ListOfUniqueCharacters.Length)].Trim();
            }
            return randomStr.Trim();
        }


        #endregion

        /// <summary>
        /// This method is used to reduce quality of the image
        /// </summary>
        /// <param name="sourceImageFilePath">
        /// full path of the source image, which needs to be compressed
        /// </param>
        /// <param name="destImageFilePath">
        /// full path of destination image, which will be saved after being compressed. 
        /// </param>
        public static void CompressImage(string sourceImageFilePath, string destImageFilePath)
        {
            try
            {
                // Get a bitmap object from source image
                Bitmap bmp = new Bitmap(sourceImageFilePath);

                // Create an Encoder object based on the GUID
                // for the Quality parameter category.
                ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);
                System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;

                // Create an EncoderParameters object.
                // An EncoderParameters object has an array of EncoderParameter objects. In this case, there is only one
                // EncoderParameter object in the array.
                EncoderParameters myEncoderParameters = new EncoderParameters(1);
                //Compression ratio for compressing the image
                long compressionRatio = Common.Property.imageCompressionRatio;


                if (Common.Property.parameterdic.ContainsKey("ScreenshotCompressionRatio"))
                {
                    float cr = float.Parse(Common.Utility.GetParameter("ScreenshotCompressionRatio"));
                    if (cr >= 0.5 && cr <= 1)
                    {
                        compressionRatio = (long)(cr * 100);
                    }
                }

                //'compressionRatio' is either default value or if present in Parameters.ini, will be used
                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, compressionRatio);
                myEncoderParameters.Param[0] = myEncoderParameter;
                bmp.Save(destImageFilePath, jgpEncoder, myEncoderParameters);
                myEncoderParameters.Dispose();
                bmp.Dispose();
            }
            catch
            {
                //If there is an exception saving the image, save original one itself
                File.Copy(sourceImageFilePath, destImageFilePath, true);
            }

        }

        public static void Savecontentstofile(string path, string contents)
        {
            try
            {
                File.WriteAllText(path, contents);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            try
            {
                ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
                foreach (ImageCodecInfo codec in codecs)
                {
                    if (codec.FormatID == format.Guid)
                    {
                        return codec;
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static void resetApp()
        {

        }

        /// <summary>
        /// Method to retrieve all TestCaseIds from given TestSuite File.
        /// </summary>
        /// <param name="suiteFilePath">string </param>
        /// <returns>string</returns>
        public static string GetTestCaseIdFromSuiteFile(string suiteFilePath)
        {
            string testCaseId = string.Empty;
            suiteFilePath = suiteFilePath + ".suite";
            string[] lines = System.IO.File.ReadAllLines(suiteFilePath);
            foreach (var id in lines)
            {
                if (id != string.Empty)
                {
                    testCaseId = testCaseId + "," + id;
                }
            }
            return testCaseId.Substring(1);
        }
    }
}
