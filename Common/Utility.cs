/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: cs
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

        public static Dictionary<int, string> DriverKeydic = new Dictionary<int, string>();

        /// <summary>
        /// Match the expected and actual values based on options specified.
        /// </summary>
        /// <param name="expectedValue"> string expected value to match.</param>
        /// <param name="actualValue"> string actual value to match.</param>
        /// <returns>bool result of match.</returns>
        public static bool DoKeywordMatch(string expectedValue, string actualValue)
        {

            if (DriverKeydic.ContainsValue("ignorespace"))
            {
                expectedValue = Regex.Replace(expectedValue, @"[\s]", String.Empty).ToLower();
                actualValue = Regex.Replace(actualValue, @"[\s]", String.Empty).ToLower();
            }
            if (DriverKeydic.ContainsValue("ignorecase"))
            {
                expectedValue = expectedValue.ToLower();
                actualValue = actualValue.ToLower();
            }
            if (!DriverKeydic.ContainsValue("partialmatch") && !DriverKeydic.ContainsValue("partmatch") &&
                DriverKeydic.ContainsValue("exactmatch"))
                return actualValue.Equals(expectedValue);
            else
                return actualValue.Contains(expectedValue);
        }

        /// <summary>
        ///  Generates a unique character string for the specified length
        /// </summary>
        /// <param name="charCount">Length of string to be generated,by default it is 10.</param>
        /// <returns>Unique string with specified length.</returns>
        public static string GenerateUniqueString(int charCount = 10)
        {
            Random random = new Random();
            string randomString = string.Empty;
            for (int i = 0; i < charCount; i++)
            {
                var randomnum = random.Next(97, 123);
                randomString = randomString + (char) randomnum;
            }
            return randomString;
        }

        /// <summary>
        /// Generate a numeral string such as zip code or mobile number
        /// </summary>
        /// <param name="charCount"> Length of number to be generated,by default it is 5,</param>
        /// <returns>Unique number with specified length in string format.</returns>
        public static string GenerateUniqueNumeral(int charCount = 5)
        {
            Random random = new Random();
            string randomString = string.Empty;
            for (int i = 0; i < charCount; i++)
            {
                var randomnum = random.Next(1, 9);
                randomString = randomString + randomnum;
            }
            return randomString;
        }

        public static string GetOsVersion()
        {
            string osName;
            try
            {
                string s = Environment.OSVersion.ToString();
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
                throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0039").Replace("{MSG}", e.Message));
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
            if (Property.Runtimedic.ContainsKey(varName))
                Property.Runtimedic[varName] = varValue;
            else
                Property.Runtimedic.Add(varName, varValue);
        }

        /// <summary>
        /// Retrives value of variable from run time dictionary. Return ZERO length string if variable
        /// could not be found in dictionary.
        /// </summary>
        /// <param name="varName"></param>
        /// <returns></returns>
        public static string GetVariable(string varName)
        {
            varName = varName.ToLower();
            if (!Property.Runtimedic.ContainsKey(varName))
                return varName;

            if (string.IsNullOrEmpty(Property.Runtimedic[varName]))
                return Property.Runtimedic[varName];


            if (varName == "firefoxprofilepath")
            {
                if (!Path.IsPathRooted(Property.Runtimedic[varName]))
                {
                    Property.Runtimedic[varName] = string.Concat(Property.IniPath, Property.Runtimedic[varName]);
                    Property.Runtimedic[varName] = Path.GetFullPath(Property.Runtimedic[varName]);
                }
            }
            return Property.Runtimedic[varName];
        }

        /// <summary>
        /// Get the parameter from parameter dictionary.
        /// </summary>
        /// <param name="parameterName">Key String for which values needed.</param>
        /// <returns>String value for given parameter.</returns>
        public static string GetParameter(string parameterName)
        {

            parameterName = parameterName.ToLower();
            if (!Property.Parameterdic.ContainsKey(parameterName))
                return parameterName;
            if (string.IsNullOrEmpty(Property.Parameterdic[parameterName]))
                return Property.Runtimedic[parameterName];
            if (parameterName == "firefoxprofilepath")
            {
                if (!Path.IsPathRooted(Property.Parameterdic[parameterName]))
                {
                    Property.Parameterdic[parameterName] = string.Concat(Property.IniPath,
                        Property.Parameterdic[parameterName]);
                    Property.Parameterdic[parameterName] = Path.GetFullPath(Property.Parameterdic[parameterName]);
                }
            }
            return Property.Parameterdic[parameterName];

        }

        /// <summary>
        /// SetParameter to Parameter dictionary.
        /// </summary>
        /// <param name="paramName">Parameter String for which value to be set.</param>
        /// <param name="paramValue">Parameter Value which will be set.</param>
        public static void SetParameter(string paramName, string paramValue)
        {
            paramName = paramName.ToLower();
            if (Property.Parameterdic.ContainsKey(paramName))
                Property.Parameterdic[paramName] = paramValue;
            else
                Property.Parameterdic.Add(paramName, paramValue);
        }

        /// <summary>
        /// Return all processes started.
        /// </summary>
        /// <returns>string array of processes.</returns>
        public static string[] GetAllProcesses()
        {
            string[] processes = new string[Property.ProcessLists.Count];
            for (int i = 0; i < Property.ProcessLists.Count; i++)
            {
                processes[i] = Property.ProcessLists[i].ToString();
            }
            return processes;
        }

        /// <summary>
        /// Set process string top global arraylist.
        /// </summary>
        /// <param name="processName">string process name to add.</param>
        public static void SetProcessParameter(string processName)
        {
            Property.ProcessLists.Add(processName);
        }

        /// <summary>
        /// Replaces variables in a string. Variables are used with a convention of {$varName}.
        /// Variable is not replaced if its values could not be found.
        /// Replace variable by {$varName} after discussion
        /// </summary>
        /// <param name="inString">Input string, in which variable needs to be replaced.</param>
        /// <returns>String with variables replaced</returns>
        public static string ReplaceVariablesInString(string inString)
        {
            var endindex = 0;
            for (int v = 0;; v++)
            {
                if (inString.Contains("{$"))
                {
                    var stindex = inString.IndexOf("{$", endindex, StringComparison.Ordinal);
                    //Break if no more variables found the string
                    if (stindex < 0)
                    {
                        break;
                    }

                    endindex = inString.IndexOf("}", stindex, StringComparison.Ordinal);
                    //Break if end } is not there in the string
                    if (endindex < 0)
                    {
                        break;
                    }
                    string keyVariable = inString.Substring(stindex + 2, (endindex - stindex - 2));
                    //Retrieve value of variable from run time dictionary
                    string value = GetVariable(keyVariable);

                    //If variable name and returned value are same, consider no variable found, and do not replace original string
                    if (value.ToLower().Equals(keyVariable.ToLower()) == false)
                    {
                        inString = inString.Replace("{$" + keyVariable + "}", value);
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
            StreamReader reportSettingsReader = null;
            StreamReader sauceLabsReader = null;
            try
            {
                FileInfo parameterFileInfo = new FileInfo(Path.Combine(Property.IniPath, Property.ParameterFileName));
                reader = new StreamReader(parameterFileInfo.FullName);
                StoreReaderContent(reader);

                // Retreiving the content of report settings in INI file
                try
                {
                    FileInfo reportSettingsFileInfo =
                        new FileInfo(Property.ApplicationPath + Property.ReportSettingsFile);
                    reportSettingsReader = new StreamReader(reportSettingsFileInfo.FullName);
                    StoreReaderContent(reportSettingsReader);
                }
                catch (Exception e)
                {
                    throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0042"), e.Message);
                }

                // Retreiving the contents of SauceLabs.ini file
                if (Property.Parameterdic[Property.BrowserString].ToLower() == Property.SauceLabs.ToLower())
                {
                    try
                    {
                        if (File.Exists(Path.Combine(Property.IniPath, Property.SauceLabsParameterFile)))
                        {
                            FileInfo sauceLabsFileInfo =
                                new FileInfo(Path.Combine(Property.IniPath, Property.SauceLabsParameterFile));
                            sauceLabsReader = new StreamReader(sauceLabsFileInfo.FullName);
                            StoreReaderContent(sauceLabsReader);
                        }
                        else
                            throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0037"), "saucelabs.ini");
                    }
                    catch (Exception e)
                    {

                        throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0059"), e.Message);
                    }
                }
                //END of sauce file access code
            }
            catch (Exception e)
            {
                throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0044"), e.Message);
            }
            finally
            {
                if (sauceLabsReader != null)
                {
                    sauceLabsReader.Close();
                }
                if (reportSettingsReader != null) reportSettingsReader.Close();
                if (reader != null) reader.Close();
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
                    FileInfo databaseFileInfo =
                        new FileInfo(Path.Combine(Property.EnvironmentFileLocation,
                            GetParameter(Property.Environment) + ".ini"));
                    using (StreamReader dbreader = new StreamReader(databaseFileInfo.FullName))
                        StoreReaderContent(dbreader);
                }
                catch (Exception e)
                {
                    throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0041"), e.Message);
                }

                // Retreiving the content of test manager specific ini file
                // This file will be named as QC.ini, MSTestManager.ini etc
                {
                    FileInfo databaseFileInfo =
                        new FileInfo(Path.Combine(Property.ApplicationPath, GetParameter("ManagerType") + ".ini"));
                    if (File.Exists(databaseFileInfo.FullName))
                    {
                        using (StreamReader dbreader = new StreamReader(databaseFileInfo.FullName))
                            StoreReaderContent(dbreader);
                    }
                }
                //Parsing any other extra ini file, that can be passed as iniFiles=TestManager.ini, Custom.ini etc.
                //This section is highly powerfull, using which you can pass on any ini file to be parsed
                string extraIniFiles = GetVariable("iniFiles");
                string[] arrIniFiles = extraIniFiles.Split(',');
                foreach (string iniFile in arrIniFiles)
                {
                    string iniFileLocation = Path.Combine(Property.ApplicationPath, iniFile.Trim());
                    if (File.Exists(iniFileLocation))
                    {
                        FileInfo databaseFileInfo = new FileInfo(iniFileLocation);
                        using (StreamReader dbreader = new StreamReader(databaseFileInfo.FullName))
                            StoreReaderContent(dbreader);
                    }
                }
            }
            catch (Exception e)
            {
                throw new KryptonException(GetCommonMsgVariable("KRYPTONERRCODE0044"), e.Message);
            }
        }

        /// <summary>
        /// Generate string (in executable arguments format ) from specified Dictionary.
        /// </summary>
        /// <param name="dicToProcess">Dictionary to process ie. Dictionary<string>
        ///         <string></string>
        ///     </string>
        /// </param>
        /// <returns>String</returns>
        public static string GenerateDictionaryContentInArgumentsFormat(Dictionary<string, string> dicToProcess)
        {
            string argumentString = string.Empty;
            try
            {
                foreach (KeyValuePair<string, string> keyValuePair in dicToProcess)
                {
                    if (argumentString.Equals(string.Empty))
                    {
                        argumentString = argumentString + "\"" + keyValuePair.Key + "=" + keyValuePair.Value + "\"";
                    }
                    else
                        argumentString = argumentString + " " + "\"" + keyValuePair.Key + "=" + keyValuePair.Value +
                                         "\"";
                }


            }
            catch (Exception)
            {
                // ignored
            }
            return argumentString;
        }

        public static string CreateTempInIfile(Dictionary<string, string> dic, Dictionary<string, string> otherDic)
        {
            string tempFilePAth = string.Empty;

            //This will avoid sinking of dictionaries.
            Dictionary<string, string> tempDic = new Dictionary<string, string>(dic);

            try
            {
                foreach (KeyValuePair<string, string> kvPair in otherDic)
                {
                    if (dic.ContainsKey(kvPair.Key))
                    {
                        tempDic[kvPair.Key] = otherDic[kvPair.Key];
                    }
                    else
                    {
                        tempDic.Add(kvPair.Key.ToLower(), kvPair.Value);
                    }
                }
                tempFilePAth = Property.ApplicationPath + "//" + Property.TempiniFileName + ".ini";
                FileStream stream = File.Create(tempFilePAth);
                StreamWriter writer = new StreamWriter(stream, Encoding.UTF8);
                // for merging mobile build
                foreach (KeyValuePair<string, string> keyValuePair in tempDic)
                {
                    writer.WriteLine(keyValuePair.Key + " : " + keyValuePair.Value);
                }
                writer.Dispose();
                stream.Dispose();
            }
            catch (Exception)
            {
                // Console.WriteLine(e.Message);
            }
            return tempFilePAth;
        }

        ///  <summary>
        /// Verify the presence of all the parameters in parameters.ini file
        /// It also checks for the values of the mandatory parameters.
        ///  </summary>
        public static void ValidateParameters()
        {
            bool validation = true;
            //    testCaseOrSuite;
            string errorMessage = string.Empty;
            Dictionary<String, int> allParameters = new Dictionary<string, int>
            {
                {"browser", 0},
                {"driver", 0},
                {"testcaselocation", 1},
                {"testdatalocation", 1},
                {"testcasefileextension", 0},
                {"reusabledefaultfile", 0},
                {"orlocation", 1},
                {"orfilename", 1},
                {"recoverfrompopuplocation", 1},
                {"recoverfrompopupfilename", 1},
                {"recoverfrombrowserlocation", 1},
                {"recoverfrombrowserfilename", 1},
                {"dbtestdatalocation", 1},
                {"logdestinationfolder", 1},
                {"testcaseid", 2},
                {"testsuite", 2},
                {"managertype", 0},
                {"timeformat", 0},
                {"dateformat", 0},
                {"datetimeformat", 0},
                {"companylogo", 0},
                {"failedcountforexit", 0},
                {"endexecutionwaitrequired", 0},
                {"dbconnectionstring", 1},
                {"defaultdb", 1},
                {"dbserver", 1},
                {"dbqueryfilepath", 1},
                {"debugmode", 0},
                {"errorcaptureas", 0},
                {"environment", 0},
                {"environmentsetupbatch", 0},
                {"testcaseidseperator", 0},
                {"testcaseidparameter", 0},
                {"snapshotoption", 0},
                {"runremoteexecution", 0},
                {"runonremotebrowserurl", 0},
                {"emailnotification", 0},
                {"emailnotificationfrom", 0},
                {"emailsmtpserver", 0},
                {"emailsmtpport", 0},
                {"emailsmtpusername", 0},
                {"emailsmtppassword", 0},
                {"emailstarttemplate", 1},
                {"emailendtemplate", 1},
                {"keepreporthistory", 0},
                {"validatesetup", 0},
                {"objecttimeout", 0},
                {"globaltimeout", 0},
                {"testmode", 0},
                {"closebrowseroncompletion", 0},
                {"firefoxprofilepath", 0},
                {"addonspath", 0},
                {"applicationurl", 1},
                {"maxtimeoutforpageload", 0},
                {"mintimeoutforpageload", 0},
                {"scriptlanguage", 0},
                {"recoverycount", 0},
                {"parallelrecoverysheetname", 1},
                {"environmentfilelocation", 1},
                {"startparallelrecovery", 2}
            };

            //Please make an entry if added a new parameter in parameters.ini file
            // verified in conjunction
            // verified in conjunction
            // allParameters.Add("Waitforalert", 0);
            try
            {
                for (int i = 0; i < allParameters.Count; i++)
                {
                    if (Property.Parameterdic.ContainsKey(allParameters.ElementAt(i).Key) == false)
                    {
                        validation = false;
                        errorMessage = GetCommonMsgVariable("KRYPTONERRCODE0045")
                            .Replace("{MSG}", allParameters.ElementAt(i).Key);
                        break;
                    }

                    if (allParameters.ElementAt(i).Value == 0) continue;
                    if (allParameters.ElementAt(i).Value == 1)
                    {
                        if (string.IsNullOrWhiteSpace(GetParameter(allParameters.ElementAt(i).Key)))
                        {
                            validation = false;
                            errorMessage = GetCommonMsgVariable("KRYPTONERRCODE0046")
                                .Replace("{MSG}", allParameters.ElementAt(i).Key);
                            break;
                        }
                    }
                    else if (allParameters.ElementAt(i).Value == 2)
                    {
                        if (string.IsNullOrWhiteSpace(GetParameter("TestCaseId")) &&
                            string.IsNullOrWhiteSpace(GetParameter("TestSuite")))
                        {
                            validation = false;
                            errorMessage = GetCommonMsgVariable("KRYPTONERRCODE0047");
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
                KryptonException.Writeexception(e);
                throw new KryptonException(e.Message);
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
                    int seperatorPos = line.IndexOf(Property.KeyValueDistinctionKeyword);
                    if (seperatorPos > 0)
                    {
                        string[] param = line.Split(Property.KeyValueDistinctionKeyword);

                        string val = string.Empty;
                        if (param.Length > 2)
                        {
                            for (int paramIndex = 1; paramIndex < param.Length; paramIndex++)
                            {
                                if (string.IsNullOrWhiteSpace(val)) val = param[paramIndex];
                                else val = val + Property.KeyValueDistinctionKeyword + param[paramIndex];
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
                            if (Property.Parameterdic.ContainsKey(param[0].ToLower().Trim()))
                            {
                                Property.Parameterdic[param[0].ToLower().Trim()] = val.Trim();
                                Property.Runtimedic[param[0].ToLower().Trim()] = val.Trim();
                            }
                            else
                            {
                                Property.Parameterdic.Add(param[0].ToLower().Trim(), val.Trim());
                                Property.Runtimedic.Add(param[0].ToLower().Trim(), val.Trim());
                            }
                        }
                        catch
                        {
                            //No throw
                        }
                    }
                }
                // set projectPath variable for using it as a variable anywhere
                if (GetParameter("projectpath").ToLower().Equals("true"))
                {
                    SetParameter("projectpath", Property.IniPath);
                    SetVariable("projectpath", Property.IniPath);
                }
            }
            finally
            {
                reader.Close();
            }
        }

        /// <summary>
        /// Initialize Browser for execution.
        /// </summary>
        /// <param name="browserrun">Browser on which execution will run.</param>
        public static void InitializeBrowser(string browserrun)
        {
            BrowserManager.GetBrowser(browserrun);
        }

        /// <summary>
        ///Email validation method to validate the format of email files.
        /// </summary>
        /// <returns>true on success</returns>
        private static bool ValidateEmail()
        {
            if (GetParameter("EmailNotification").Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                if (GetParameter("EmailNotificationFrom") == string.Empty ||
                    !GetParameter("EmailNotificationFrom").Contains("@"))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_NOTIFICATION_ABORTED);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPServer") == string.Empty || !GetParameter("EmailSMTPServer").Contains("."))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_NOTIFICATION_ABORTED);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                long num;
                if (GetParameter("EmailSMTPPort") == string.Empty ||
                    !long.TryParse(GetParameter("EmailSMTPPort"), out num))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_NOTIFICATION_ABORTED);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPUsername") == string.Empty ||
                    !GetParameter("EmailSMTPUsername").Contains("@"))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_NOTIFICATION_ABORTED);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                if (GetParameter("EmailSMTPPassword") == string.Empty)
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_NOTIFICATION_ABORTED);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
                // so that templates folder will accept both absolute and relative paths.
                string startTemplatePath = Property.EmailStartTemplate;
                if (!File.Exists(startTemplatePath))
                {
                    startTemplatePath = Property.ApplicationPath + GetParameter("EmailStartTemplate");
                }
                if (!File.Exists(startTemplatePath))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_TEMPLETE_MISSING);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }

                string endTemplatePath = Property.EmailEndTemplate;
                if (!File.Exists(endTemplatePath))
                {
                    endTemplatePath = Property.ApplicationPath + GetParameter("EmailEndTemplate");
                }
                if (!File.Exists(endTemplatePath))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_TEMPLETE_MISSING);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
                string startTemplate = File.ReadAllText(startTemplatePath);
                if (!startTemplate.ToLower().Contains("subject::") || !startTemplate.ToLower().Contains("body::"))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_TEMPLETE_MISSING);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
                string endTemplate = File.ReadAllText(endTemplatePath);
                if (!endTemplate.ToLower().Contains("subject::") || !endTemplate.ToLower().Contains("body::"))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_TEMPLETE_MISSING);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
                string emailNotif = File.ReadAllText(Path.Combine(Property.IniPath, "EmailNotification.ini"));
                if (!emailNotif.Trim().ToLower().Contains("recipient"))
                {
                    Console.WriteLine(ConsoleMessages.EMAIL_TEMPLETE_MISSING);
                    Console.WriteLine(ConsoleMessages.MSG_DASHED);
                    return false;
                }
            }
            return true;
        }

        ///  <summary>
        /// Email Notification method for different stages
        ///  </summary>
        ///  <param name="stage">stage/state of execution</param>
        /// <param name="zipAttch"></param>
        /// <param name="reportFileName"></param>
        public static void EmailNotificationOld(string stage, bool zipAttch, string reportFileName = null)
        {
            if (ValidateEmail() == false) return;

            Console.WriteLine(ConsoleMessages.MSG_DASHED);
            Console.WriteLine("Sending Email....");
            Console.WriteLine(ConsoleMessages.MSG_DASHED);

            try
            {
                string[] emailStr = null;
                switch (stage.ToLower())
                {

                    case "start":
                        emailStr = EmailTemplate(Property.EmailStartTemplate);
                        break;
                    case "end":
                        emailStr = EmailTemplate(Property.EmailEndTemplate);
                        break;
                }

                MailMessage message = new MailMessage();

                MailAddress mailFrom = new MailAddress(GetParameter("EmailNotificationFrom"), "Krypton Automation");

                message.From = mailFrom;

                if (emailStr != null)
                {
                    var mailSubject = emailStr[1];
                    mailSubject = mailSubject.Replace("[Test Suite]",
                        string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false
                            ? GetParameter("TestSuite")
                            : string.Empty);
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
                    SmtpClient smtpClient = new SmtpClient(GetParameter("EmailSMTPServer"),
                        int.Parse(GetParameter("EmailSMTPPort")));
                    NetworkCredential credential = new NetworkCredential(GetParameter("EmailSMTPUsername"),
                        GetParameter("EmailSMTPPassword"));
                    smtpClient.Credentials = credential;
                    smtpClient.EnableSsl = true;
                    if (File.Exists(Path.Combine(Property.IniPath, Property.EmailNotificationFile)))
                    {
                        try
                        {
                            string allrecipients =
                                File.ReadAllText(Path.Combine(Property.IniPath, Property.EmailNotificationFile));
                            allrecipients = allrecipients.Substring(allrecipients.IndexOf(':') + 1);
                            string[] recipient = allrecipients.Split(',');
                            {
                                foreach (string recipientAdd in recipient)
                                {
                                    if ((!string.IsNullOrWhiteSpace(recipientAdd)) &&
                                        (!string.IsNullOrEmpty(recipientAdd)))
                                        message.To.Add(new MailAddress(recipientAdd.Trim(), null));
                                }
                            }
                        }
                        catch (Exception)
                        {
                            throw new Exception("Invalid Email format in EmailNotification.ini");
                        }

                    }
                    if (string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false)
                        emailStr[3] = emailStr[3].Replace("[Test Suite]", GetParameter("TestSuite"));
                    else
                        emailStr[3] = emailStr[3].Replace("[Test Suite]", string.Empty);
                    emailStr[3] = emailStr[3].Replace("[Result]", Property.FinalExecutionStatus);
                    emailStr[3] = emailStr[3].Replace("[StartTime]", Property.ExecutionStartDateTime);
                    if (stage == "end")
                    {
                        emailStr[3] = emailStr[3].Replace("[EndTime]", Property.ExecutionEndDateTime);

                        DateTime executionStartTime = DateTime.ParseExact(Property.ExecutionStartDateTime,
                            GetParameter("DateTimeFormat"), null);
                        DateTime executionEndTime = DateTime.ParseExact(Property.ExecutionEndDateTime,
                            GetParameter("DateTimeFormat"), null);
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
                    emailStr[3] = emailStr[3].Replace("[LogDestinationFolder]",
                        Property.ResultsDestinationPath.Substring(0, Property.ResultsDestinationPath.LastIndexOf('\\')));


                    if (stage == "end")
                    {
                        // Create  the file attachment for this e-mail message.
                        string file = Property.HtmlFileLocation + "/" + Property.ReportZipFileName;

                        int doStep = 0;
                        do
                        {
                            Attachment data;
                            if (File.Exists(file) && zipAttch)
                            {
                                FileInfo fileInfo = new FileInfo(file);
                                long fileSize = fileInfo.Length;
                                if (fileSize > 1024*1024*4)
                                {
                                    emailStr[3] = emailStr[3] +
                                                  "\n---------------\n**NOTE: Not able to send attachement, size of the attachment is too big.**";

                                    string htmlFile = Property.HtmlFileLocation + "/HtmlReport.html";
                                    data = new Attachment(htmlFile, MediaTypeNames.Application.Octet);
                                    // Add time stamp information for the file.
                                    ContentDisposition htmlDisposition = data.ContentDisposition;
                                    htmlDisposition.CreationDate = File.GetCreationTime(htmlFile);
                                    htmlDisposition.ModificationDate = File.GetLastWriteTime(htmlFile);
                                    htmlDisposition.ReadDate = File.GetLastAccessTime(htmlFile);
                                    // Add the html file attachment to this e-mail message.
                                    message.Attachments.Add(data);

                                    break;
                                }

                                data = new Attachment(file, MediaTypeNames.Application.Octet);
                                // Add time stamp information for the file.
                                ContentDisposition disposition = data.ContentDisposition;
                                disposition.CreationDate = File.GetCreationTime(file);
                                disposition.ModificationDate = File.GetLastWriteTime(file);
                                disposition.ReadDate = File.GetLastAccessTime(file);
                                // Add the zip file attachment to this e-mail message.
                                message.Attachments.Add(data);

                                break;
                            }
                            if (zipAttch == false)
                            {
                                string htmlFile;
                                if (string.IsNullOrWhiteSpace(reportFileName))
                                {
                                    htmlFile = Property.HtmlFileLocation + "/HtmlReportMail.html";

                                }
                                else
                                {
                                    htmlFile = reportFileName;
                                }

                                data = new Attachment(htmlFile, MediaTypeNames.Application.Octet);
                                // Add time stamp information for the file.
                                ContentDisposition htmlDisposition = data.ContentDisposition;
                                htmlDisposition.CreationDate = File.GetCreationTime(htmlFile);
                                htmlDisposition.ModificationDate = File.GetLastWriteTime(htmlFile);
                                htmlDisposition.ReadDate = File.GetLastAccessTime(htmlFile);
                                // Add the html file attachment to this e-mail message.
                                message.Attachments.Add(data);

                                break;
                            }
                            doStep++;
                            if (doStep < 5)
                                Thread.Sleep(5000); //wait for 5 sec.
                        } while (doStep < 5);

                    }

                    message.Subject = mailSubject;
                    if (!string.IsNullOrEmpty(Property.ReportSummaryBody))
                    {
                        message.Body = Property.ReportSummaryBody;
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

            Console.WriteLine(ConsoleMessages.MSG_DASHED);
            Console.WriteLine("Sending Email ....!!!");
            Console.WriteLine(ConsoleMessages.MSG_DASHED);

            try
            {
                string[] emailStr = null;
                switch (stage.ToLower())
                {

                    case "start":
                        emailStr = EmailTemplate(Property.EmailStartTemplate);
                        break;
                    case "end":
                        emailStr = EmailTemplate(Property.EmailEndTemplate);
                        break;
                }


                #region Email Configuration settings

                CDO.Message message = new CDO.Message();
                CDO.IConfiguration configuration = message.Configuration;
                ADODB.Fields fields = configuration.Fields;
                ADODB.Field field = fields["http://schemas.microsoft.com/cdo/configuration/smtpserver"];
                field.Value = GetParameter("EmailSMTPServer").Trim();

                field = fields["http://schemas.microsoft.com/cdo/configuration/smtpserverport"];
                field.Value = int.Parse(GetParameter("EmailSMTPPort").Trim());

                field = fields["http://schemas.microsoft.com/cdo/configuration/sendusing"];
                field.Value = CDO.CdoSendUsing.cdoSendUsingPort;

                field = fields["http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"];
                field.Value = CDO.CdoProtocolsAuthentication.cdoBasic;

                field = fields["http://schemas.microsoft.com/cdo/configuration/sendusername"];
                field.Value = GetParameter("EmailSMTPUsername").Trim();

                field = fields["http://schemas.microsoft.com/cdo/configuration/sendpassword"];
                field.Value = GetParameter("EmailSMTPPassword").Trim();

                field = fields["http://schemas.microsoft.com/cdo/configuration/smtpusessl"];
                field.Value = "true";
                fields.Update();

                #endregion



                #region  Add attachments

                if (stage == "end")
                {
                    // Create  the file attachment for this e-mail message.
                    string file = Property.HtmlFileLocation + "/" + Property.ReportZipFileName;
                    if (File.Exists(file) && zipAttch)
                    {
                        message.AddAttachment(file);

                    }
                    else if (!zipAttch)
                    {
                        string htmlFile;
                        if (string.IsNullOrWhiteSpace(reportFileName))
                        {
                            htmlFile = Property.HtmlFileLocation + "/HtmlReportMail.html";
                        }
                        else
                        {
                            htmlFile = reportFileName;
                        }
                        message.AddAttachment(htmlFile);
                    }

                }

                #endregion

                if (!string.IsNullOrEmpty(Property.ReportSummaryBody))
                {
                    message.HTMLBody = Property.ReportSummaryBody;

                }
                else
                {
                    if (emailStr != null)
                    {
                        string mailTextBody = GetEmailTextBody(emailStr[3], stage);
                        message.TextBody = mailTextBody;
                    }
                }
                message.From = "KryptonAutomation";
                message.Sender = GetParameter("EmailSMTPUsername").Trim();
                message.To = GetRecipientList();
                if (emailStr != null) message.Subject = GetEmailSubjectMessage(emailStr[1]);
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
            var EmailTextBody = emailTextBody.Replace("[LogDestinationFolder]",
                Property.ResultsDestinationPath.Substring(0, Property.ResultsDestinationPath.LastIndexOf('\\')));
            try
            {
                emailTextBody = emailTextBody.Replace("[Test Suite]",
                    string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false
                        ? GetParameter("TestSuite")
                        : string.Empty);
                emailTextBody = emailTextBody.Replace("[Result]", Property.FinalExecutionStatus);
                emailTextBody = emailTextBody.Replace("[StartTime]", Property.ExecutionStartDateTime);
                if (stage == "end")
                {
                    emailTextBody = emailTextBody.Replace("[EndTime]", Property.ExecutionEndDateTime);

                    DateTime executionStartTime = DateTime.ParseExact(Property.ExecutionStartDateTime,
                        GetParameter("DateTimeFormat"), null);
                    DateTime executionEndTime = DateTime.ParseExact(Property.ExecutionEndDateTime,
                        GetParameter("DateTimeFormat"), null);
                    TimeSpan time = executionEndTime - executionStartTime;

                    emailTextBody = emailTextBody.Replace("[ExecutionTime]", time.ToString());
                }
                emailTextBody = emailTextBody.Replace("[RemoteUrl]",
                    GetParameter("RunRemoteExecution") == "true" ? GetParameter("RunOnRemoteBrowserUrl") : "localhost");

                emailTextBody = emailTextBody.Replace("[Environment]", GetParameter("Environment"));

                emailTextBody = emailTextBody.Replace("[TotalStepExecuted]", Property.TotalStepExecuted.ToString());
                emailTextBody = emailTextBody.Replace("[TotalStepPass]", Property.TotalStepPass.ToString());
                emailTextBody = emailTextBody.Replace("[TotalStepFail]", Property.TotalStepFail.ToString());
                emailTextBody = emailTextBody.Replace("[TotalStepWarning]", Property.TotalStepWarning.ToString());

                emailTextBody = emailTextBody.Replace("[TotalCaseExecuted]", Property.TotalCaseExecuted.ToString());
                emailTextBody = Property.FinalExecutionStatus.ToLower().Equals("passed")
                    ? emailTextBody.Replace("[TotalCaseFail]/", string.Empty)
                    : emailTextBody.Replace("[TotalCaseFail]", Property.TotalCaseFail.ToString());

                emailTextBody = emailTextBody.Replace("[TotalCasePass]", Property.TotalCasePass.ToString());
                // ReSharper disable once ReturnValueOfPureMethodIsNotUsed
                emailTextBody.Replace("[TotalCaseWarning]", Property.TotalCaseWarning.ToString());

                //Results Destination Path variable
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
        /// <param name="emailMessage"></param>
        /// <returns></returns>
        private static string GetEmailSubjectMessage(string emailMessage)
        {
            try
            {
                var mailSubject = emailMessage;
                mailSubject = mailSubject.Replace("[Test Suite]",
                    string.IsNullOrWhiteSpace(GetParameter("TestSuite")) == false ? GetParameter("TestSuite") : "");
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
                mailSubject = Property.FinalExecutionStatus.ToLower().Equals("passed")
                    ? mailSubject.Replace("[TotalCaseFail]/", string.Empty)
                    : mailSubject.Replace("[TotalCaseFail]", Property.TotalCaseFail.ToString());
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
            string recipientList = string.Empty;
            if (File.Exists(Path.Combine(Property.IniPath, Property.EmailNotificationFile)))
            {
                try
                {
                    string allrecipients =
                        File.ReadAllText(Path.Combine(Property.IniPath, Property.EmailNotificationFile));
                    allrecipients = allrecipients.Substring(allrecipients.IndexOf(':') + 1);
                    string[] recipient = allrecipients.Split(',');
                    {
                        recipientList =
                            recipient.Where(
                                recipientAdd =>
                                    (!string.IsNullOrWhiteSpace(recipientAdd)) && (!string.IsNullOrEmpty(recipientAdd)))
                                .Aggregate(recipientList, (current, recipientAdd) => current + "," + recipientAdd.Trim());
                    }
                }
                catch (Exception)
                {
                    throw new Exception("Invalid Email format in EmailNotification.ini");
                }
            }
            return recipientList;
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
                for (int v = 0;; v++)
                {
                    if (optionValue.Contains("{"))
                    {
                        int stindex = optionValue.IndexOf("{", StringComparison.Ordinal);
                        optionValue = optionValue.Remove(stindex, 1);

                        int endindex = optionValue.IndexOf("}", StringComparison.Ordinal);
                        if (endindex < 0) break;

                        string keyVariable = optionValue.Substring(stindex, (endindex - stindex));
                        // Modified to check the existence of 'script' or 'testmode' keywords before '=' only.
                        if (keyVariable.IndexOf(optionType, StringComparison.OrdinalIgnoreCase) >= 0 &&
                            keyVariable.Contains("="))
                            if (keyVariable.IndexOf(optionType, StringComparison.OrdinalIgnoreCase) <
                                keyVariable.IndexOf("=", StringComparison.Ordinal))
                                testModeStr = keyVariable.Split('=')[0].Trim().Replace("\"", "") + "," +
                                              keyVariable.Split('=')[1].Trim().Replace("\"", "");
                        stindex = optionValue.IndexOf("}", StringComparison.Ordinal);
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
        public static Dictionary<string, string> GetTestOrData(string parent, string testObj, DataSet orTestData)
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
            foreach (
                DataRow drData in
                    orTestData.Tables[0].Rows.Cast<DataRow>()
                        .Where(drData => drData["parent"].ToString().Equals(parent, StringComparison.OrdinalIgnoreCase)
                                         &&
                                         drData[KryptonConstants.TEST_OBJECT].ToString()
                                             .Equals(testObj, StringComparison.OrdinalIgnoreCase)))
            {
                if (orDataRow.ContainsKey("logical_name") == false)
                {
                    orDataRow.Add("parent", ReplaceVariablesInString(drData["parent"].ToString().Trim()));
                    orDataRow.Add("logical_name", ReplaceVariablesInString(drData["logical_name"].ToString().Trim()));
                    orDataRow.Add(KryptonConstants.OBJ_TYPE,
                        ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString().Trim()));
                    orDataRow.Add(KryptonConstants.HOW,
                        ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString().Trim()));
                    orDataRow.Add(KryptonConstants.WHAT,
                        ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString().Trim()));
                    orDataRow.Add(KryptonConstants.MAPPING,
                        ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString().Trim()));
                }
                else
                {
                    orDataRow.Add("parent", ReplaceVariablesInString(drData["parent"].ToString().Trim()));
                    orDataRow["logical_name"] = ReplaceVariablesInString(drData["logical_name"].ToString());
                    orDataRow[KryptonConstants.OBJ_TYPE] =
                        ReplaceVariablesInString(drData[KryptonConstants.OBJ_TYPE].ToString());
                    orDataRow[KryptonConstants.HOW] = ReplaceVariablesInString(drData[KryptonConstants.HOW].ToString());
                    orDataRow[KryptonConstants.WHAT] = ReplaceVariablesInString(drData[KryptonConstants.WHAT].ToString());
                    orDataRow[KryptonConstants.MAPPING] =
                        ReplaceVariablesInString(drData[KryptonConstants.MAPPING].ToString());
                }
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
                {
                    /*leave the locked files as they are not required for profile*/
                }
            }

            //find all the directories into source directory and make a recursive call to copy its file to destination folder.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }

            return true;
        }

        /// <summary>
        ///Method to make a backup of firefox profile after execution of test case completed.
        ///This is used to keep the cookies stored after completion of test case.
        /// </summary>
        public static void FirefoxBackup()
        {
            if (GetParameter("FirefoxProfilePath").Length != 0 &&
                GetParameter("Browser").Equals("firefox", StringComparison.OrdinalIgnoreCase))
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
        public static void FirefoxCleanup()
        {
            if (GetParameter("FirefoxProfilePath").Length != 0 &&
                GetParameter("Browser").Equals("firefox", StringComparison.OrdinalIgnoreCase))
            {
                //we are reverting back to b3
                if (Directory.Exists(GetParameter("tmpFFProfileDir")))
                    Directory.Delete(GetParameter("tmpFFProfileDir"), true);
            }

            if (GetParameter("AddonsPath").Length != 0 &&
                GetParameter("Browser").Equals("firefox", StringComparison.OrdinalIgnoreCase))
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


            for (int v = 0;; v++)
            {
                if (optionValue.Contains("{"))
                {
                    int stindex = optionValue.IndexOf("{", StringComparison.Ordinal);
                    if (stindex > -1)
                    {
                        optionValue = optionValue.Remove(stindex, 1);
                        int endindex = optionValue.IndexOf("}", StringComparison.Ordinal);
                        if (endindex > -1)
                        {
                            string keyVariable = optionValue.Substring(stindex, (endindex - stindex));
                            if (keyVariable.IndexOf(optionType, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                testModeStr = keyVariable;
                            }
                            stindex = optionValue.IndexOf("}", StringComparison.Ordinal);
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
            Property.Runtimedic.Clear();
            Property.Runtimedic = new Dictionary<string, string>(Property.Parameterdic);
        }

        /// <summary>
        ///Creating temporary file
        /// </summary>
        /// <param name="extn"></param>
        /// <returns></returns>
        public static string GetTemporaryFile(string extn)
        {
            if (!extn.StartsWith("."))
                extn = "." + extn;

            var response = Path.GetTempPath() + Guid.NewGuid() + extn;

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
            string oriFilePath = Property.ApplicationPath + Property.GlobalCommentsFile;
            var tmpFileName = GetTemporaryFile(".txt");
            if (File.Exists(oriFilePath))
            {
                File.Copy(oriFilePath, tmpFileName);

                if (File.Exists(tmpFileName))
                {
                    using (
                        FileStream fileStream = new FileStream(tmpFileName, FileMode.Open, FileAccess.Read,
                            FileShare.None))
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        string readLine;
                        while ((readLine = streamReader.ReadLine()) != null)
                        {
                            int splitInd = readLine.IndexOf('|');
                            string messageId = readLine.Substring(0, splitInd);
                            string messageStr = readLine.Substring(splitInd + 1, readLine.Length - splitInd - 1);
                            SetCommonMsgVariable(messageId, messageStr);
                        }
                    }
                    Property.ListOfFilesInTempFolder.Add(tmpFileName);
                }
                else
                {
                    throw new KryptonException("Error", "File missing:" + tmpFileName);
                }
            }
            else
            {
                throw new KryptonException("Error", "File missing:" + oriFilePath);
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
            if (Property.CommonMsgdic.ContainsKey(varName))
                Property.CommonMsgdic[varName] = varValue;
            else
                Property.CommonMsgdic.Add(varName, varValue);
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
            string value;
            try
            {
                value = Property.CommonMsgdic[varName];
            }
            catch (Exception)
            {
                return varName; //return variable name itself if given varaible has no value.
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

        private static string[] GetRegistryKey(string softwareKey)
        {
            RegistryKey rk = Registry.LocalMachine.OpenSubKey(softwareKey);
            return rk.GetSubKeyNames();
        }

        /// <summary>
        /// Get Sub Key String For firefox browser.
        /// </summary>
        /// <returns>string : subKey String.</returns>
        public static string GetFfVersionString()
        {
            string softwareKey = Property.BrowserKeyString;
            string versionString = string.Empty;
            try
            {
                string[] subKeys = GetRegistryKey(softwareKey);
                foreach (string key in subKeys)
                {
                    if (key.Contains("Mozilla Firefox"))
                        versionString = key;
                }
                //For Win 7 machines.
                if (versionString.Equals(String.Empty))
                {
                    string[] keys = GetRegistryKey(Property.BrowserKeyString64Bit);
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
        public static void RegisterDll(string dllPath)
        {
            //registering dll silentltly.
            string fileinfo = "/s" + " " + "\"" + dllPath + "\"";
            try
            {
                Process reg = new Process
                {
                    StartInfo =
                    {
                        FileName = "regsvr32.exe",
                        Arguments = fileinfo,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true
                    }
                };
                reg.Start();
                reg.WaitForExit();
                reg.Close();
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// Method to retrieve all numbers from given string.
        /// </summary>
        /// <param name="text">string </param>
        /// <returns>number</returns>
        public static int GetNumberFromText(string text)
        {
            string number =
                (from c in text let a = Convert.ToInt32(c) where a >= 48 && a <= 57 select c).Aggregate(string.Empty,
                    (current, c) => current + c.ToString());

            return Int32.Parse(number);
        }

        #region random number generation

        /// <summary>
        ///These steps need to be written to have syncronize random number generation
        /// so that number will be trully random
        /// </summary>

        private static readonly Random Random = new Random();

        private static readonly object SyncLock = new object();

        public static int RandomNumber(int min, int max)
        {
            lock (SyncLock)
            {
                // synchronize
                return Random.Next(min, max);
            }
        }

        public static string RandomString(int strLength)
        {
            string randomStr = string.Empty;
            for (int chrCnt = 0; chrCnt < strLength; chrCnt++)
            {
                randomStr = randomStr +
                            Property.ListOfUniqueCharacters[RandomNumber(0, Property.ListOfUniqueCharacters.Length)]
                                .Trim();
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
                long compressionRatio = Property.ImageCompressionRatio;


                if (Property.Parameterdic.ContainsKey("ScreenshotCompressionRatio"))
                {
                    float cr = float.Parse(GetParameter("ScreenshotCompressionRatio"));
                    if (cr >= 0.5 && cr <= 1)
                    {
                        compressionRatio = (long) (cr*100);
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
                return codecs.FirstOrDefault(codec => codec.FormatID == format.Guid);
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
        /// <param name="suiteDataSet">DataSet</param>
        /// <returns>string</returns>
        public static string GetTestCaseIdFromSuiteDataset(DataSet suiteDataSet)
        {
            StringBuilder testcaseId = new StringBuilder();
            List<object> testcaseIdList =
                suiteDataSet.Tables[0].AsEnumerable()
                    .Select(r => r["test_case_id"])
                    .Where(x => x.ToString() != "")
                    .ToList();
            List<object> optionsList =
                suiteDataSet.Tables[0].AsEnumerable().Select(opt => opt["options"]).Take(testcaseIdList.Count).ToList();
            if (testcaseIdList[0].ToString().Contains("To enable"))
            {
                testcaseIdList.RemoveAt(0);
                optionsList.RemoveAt(0);   
            }
            for (int i = 0; i < testcaseIdList.Count; i++)
            {
                if (optionsList[i].ToString() != "{skip}")
                {
                    testcaseId.Append("," + testcaseIdList[i]);
                }
            }
            return testcaseId.ToString().Substring(1);
        }
    }
}
