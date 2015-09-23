/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: BrowserManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software 
** Description: Handle Windows PopUp Recovery 
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Excel;
using System.Diagnostics;
using AutoItX3Lib;
using System.Runtime.Serialization.Formatters.Binary;


namespace KRYPTONParallelRecovery
{
    class Program
    {
        #region Static Member Variables.
        static AutoItX3Lib.AutoItX3 objAutoit=null;
        static string sheetName = string.Empty;
        static string applicationPath = string.Empty;
        static string parentProcessId = null;
        static bool isParentProcessRunning = true;
        #endregion
        #region Main Execution
        /// <summary>
        /// Execution starts here
        /// </summary>
        /// <param name="args"> Two agruments would be passed 1. Recovery file path 2. Sheet name for popup file. 3. Invoking Parent Process ID </param>
        static void Main(string[] args)
        {
            try
            {
                string filePath = args[0].ToString();
                Console.WriteLine("FilePath" + filePath);
                applicationPath = Application.ExecutablePath;
                dynamic windowsList = null;
                List<string> windowTitleList = null;
                int applicationFilePath = applicationPath.LastIndexOf("\\");
                string WorkingDirectory = string.Empty;
                Process parentProcess = null;
                if (applicationFilePath >= 0)
                {
                    WorkingDirectory = applicationPath.Substring(0, applicationFilePath + 1);
                }

                DataSet dataset = GetDataSet(filePath);

                try
                {
                    //Generate Autoit object.
                    objAutoit = new AutoItX3Lib.AutoItX3();
                    //Setting AutoIt Title match to Exact match.
                    objAutoit.AutoItSetOption("WinTitleMatchMode", 4);
                    parentProcessId = args[2].ToString();
                }
                catch (Exception e)
                {
                    Common.KryptonException.writeexception(e);
                }

                //Starting Infinite loop.
                //loop should run as long as parent process is running
                while (isParentProcessRunning)
                {
                    //Getting all windows beforehand.
                    windowsList = objAutoit.WinList("[REGEXPTITLE:^.*$]");
                    int winListCount = (int)windowsList[0, 0]; // List[0,0] give the count of windows.

                    windowTitleList = new List<string>(); // Declaring list that would contain titles of all window.
                    //Populating non empty titles to windowTitleList.
                    for (int j = 1; j <= winListCount; j++)
                    {
                        string title = (string)windowsList[0, j];
                        if (!string.IsNullOrEmpty(title))
                        {
                            windowTitleList.Add(title);
                        }
                    }
                    if (dataset.Tables.Count > 0)
                    {
                        for (int i = 0; i < dataset.Tables[0].Rows.Count; i++)
                        {
                            //First thing first, check if parent process is running
                            //If parent process (Krypton.exe) is not running, exits from here
                            try
                            {
                                parentProcess = Process.GetProcessById(int.Parse(parentProcessId));
                            }
                            catch (Exception e)
                            {
                                //Exception will occur only when parent process is not running
                                isParentProcessRunning = false;
                                return;
                            }

                            //Get IE popup Title.
                            string popupName = dataset.Tables[0].Rows[i]["PopUpTitle"].ToString().Trim();
                            //Get IE popup Title.
                            string buttonName = dataset.Tables[0].Rows[i]["ButtonName"].ToString().Trim();
                            //Get the action need to be done.
                            string action = dataset.Tables[0].Rows[i]["Action"].ToString().Trim();

                            string data = string.Empty;
                             //Populate data if provided.
                            data = dataset.Tables[0].Rows[i]["Data"].ToString().Trim();
                            if (windowTitleList.Contains(popupName))
                            {
                                //Switching control based on action given.
                                switch (action.ToLower())
                                {
                                    case "click":
                                        btnClick(popupName, buttonName);
                                        break;
                                    case "keypress":
                                        keyPress(popupName, data);
                                        break;
                                    case "close":
                                        objAutoit.WinClose(popupName);
                                        break;
                                    case "run":
                                    case "execute":
                                        if (!File.Exists(applicationPath + data))
                                        {
                                            data = data + ".exe";
                                        }
                                         //Create process information and assign data
                                        using (Process specialScriptProcess = new Process())
                                        {
                                            specialScriptProcess.StartInfo.CreateNoWindow = false;
                                            specialScriptProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                                            specialScriptProcess.StartInfo.UseShellExecute = false;
                                            specialScriptProcess.StartInfo.FileName = data;
                                            specialScriptProcess.StartInfo.WorkingDirectory = WorkingDirectory;
                                            specialScriptProcess.StartInfo.ErrorDialog = false;
                                            // Start the process
                                            specialScriptProcess.Start();
                                            specialScriptProcess.WaitForExit(10000);
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }
                            System.Threading.Thread.Sleep(10);
                        }
                        System.Threading.Thread.Sleep(2000);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Please register AutoIt dll");
            }
            catch (System.IndexOutOfRangeException)
            {
                Console.WriteLine("KRYPTONPrallelRecovery : invalid arguments");
            }
            catch (Exception e)
            {
                Console.WriteLine("KRYPTONPrallelRecovery : " + e.Message + e.StackTrace);
            }

        }
        #endregion
        #region Private Static Methods.
        /// <summary>
        /// Key Press -> Send specified Key to window.
        /// </summary>
        /// <param name="popUpName">string : Name of popup to recover.</param>
        /// <param name="key"> string : Key to send.</param>
        private static void keyPress(string popUpName, string key)
        {
            try
            {
                objAutoit.WinActivate(popUpName, string.Empty); //Activate window.

                switch (key.ToLower())//Send Specified Key.
                {
                    case "tab":
                        objAutoit.Send("{TAB}");
                        break;
                    case "space":
                        objAutoit.Send("{SPACE}");
                        break;
                    case "enter":
                        objAutoit.Send("{ENTER}");
                        break;
                    default:
                        objAutoit.Send(key, 1);
                        break;
                }

            }
            catch (Exception e) 
            {
                Common.KryptonException.writeexception(e.InnerException);
            }
        }

        /// <summary>
        /// btnClick : Click on any object in specified window.
        /// </summary>
        /// <param name="popUpName">string : name of popup to recover.</param>
        /// <param name="btnTitle"></param>
        private static void btnClick(string popUpName, string btnTitle = "")
        {
            try
            {
                objAutoit.WinActivate(popUpName);

                //Click on control, close dialog if unsuccessful
                if (!(btnTitle.Equals("") || btnTitle.Equals(string.Empty)))
                {
                    if (objAutoit.ControlClick(popUpName, "", "[TEXT:" + btnTitle + "]") == 0)
                    {
                    }
                }
            }
            catch (Exception e)
            {
                Common.KryptonException.writeexception(e.InnerException);
            }
        }

        /// <summary>
        ///   Get Dataset from the file specified.
        /// </summary>
        /// <param name="filePath">string : Filepath.</param>
        /// <returns>Dataset : dataset for the file.</returns>
        /// 
        private static DataSet GetDataSet(string filePath)
        {

            DataSet dataset = new DataSet();
            string tmpFileName = string.Empty;
            string fileExtension = Path.GetExtension(filePath);

            try
            {
                tmpFileName = GetTemporaryFile(fileExtension, filePath);
                if (fileExtension.ToLower().Equals(".csv"))
                {
                    if (!File.Exists(tmpFileName))
                    {
                        Console.WriteLine("File Not Found:  " + filePath);
                    }
                    using (GenericParsing.GenericParserAdapter gp = new GenericParsing.GenericParserAdapter(tmpFileName, Encoding.UTF7))
                    {
                        gp.FirstRowHasHeader = true;
                        gp.ColumnDelimiter = ',';
                        dataset = gp.GetDataSet();
                    }
                }
                else
                {
                    dataset = GetExcelDataSet(filePath);
                }
            }
            catch (Exception ex) 
            {
                Common.KryptonException.writeexception(ex);
            }
            finally
            {
                File.Delete(tmpFileName);
            }
            return dataset;
        }


        private static DataSet GetExcelDataSet(string ExcelFilePath)
        {
            string tempFilePath = Path.GetTempFileName();
            IExcelDataReader excelReader = null;
            DataSet oResultSet = null;
            FileInfo oFileInfo = new FileInfo(ExcelFilePath);
            File.Copy(oFileInfo.FullName, tempFilePath, true);
            FileStream stream = File.OpenRead(tempFilePath);
            try
            {
                if (oFileInfo.Extension.ToLower() == ".xls")
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                else
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                excelReader.IsFirstRowAsColumnNames = true;
                oResultSet = excelReader.AsDataSet();

                File.Delete(tempFilePath);
            }
            catch
            {
                oResultSet = null;
                throw;
            }
            finally
            {

                if (excelReader != null)
                    excelReader.Close();
                if (stream != null)
                    stream.Close();
            }
            return oResultSet;
        }

        /// <summary>
        ///Get Temp file specified.
        /// </summary>
        /// <param name="filePath" file extension>string : fileExtension.</param>
        /// /// <param name="filePath" >string : Filepath.</param>
        /// <returns>Temp file.</returns>
        private static string GetTemporaryFile(string fileExtension, string origPath)
        {
            string response = string.Empty;
            try
            {
                if (!fileExtension.StartsWith("."))
                    fileExtension = "." + fileExtension;
                response = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + fileExtension;
                File.Copy(origPath, response);
            }
            catch
            {

            }
            return response;
        }
        #endregion
    }
}
