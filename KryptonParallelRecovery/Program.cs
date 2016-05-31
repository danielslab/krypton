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
using System.Text;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Common;
using Excel;


namespace KRYPTONParallelRecovery
{
    class Program
    {
        #region Static Member Variables.
        static AutoItX3Lib.AutoItX3 _objAutoit=null;
        static string _sheetName = string.Empty;
        static string _applicationPath = string.Empty;
        static string _parentProcessId = null;
        static bool _isParentProcessRunning = true;
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
                _applicationPath = Application.ExecutablePath;
                int applicationFilePath = _applicationPath.LastIndexOf("\\", StringComparison.Ordinal);
                string workingDirectory = string.Empty;
                if (applicationFilePath >= 0)
                {
                    workingDirectory = _applicationPath.Substring(0, applicationFilePath + 1);
                }
                DataSet dataset = GetDataSet(filePath);
                try
                {
                    //Generate Autoit object.
                    _objAutoit = new AutoItX3Lib.AutoItX3();
                    //Setting AutoIt Title match to Exact match.
                    _objAutoit.AutoItSetOption("WinTitleMatchMode", 4);
                    _parentProcessId = args[2];
                }
                catch (Exception e)
                {
                    KryptonException.Writeexception(e);
                }

                //Starting Infinite loop.
                //loop should run as long as parent process is running
                while (_isParentProcessRunning)
                {
                    //Getting all windows beforehand.
                    dynamic windowsList = _objAutoit.WinList("[REGEXPTITLE:^.*$]");
                    int winListCount = (int)windowsList[0, 0]; // List[0,0] give the count of windows.

                    var windowTitleList = new List<string>();
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
                                Process.GetProcessById(int.Parse(_parentProcessId));
                            }
                            catch (Exception e)
                            {
                                //Exception will occur only when parent process is not running
                                _isParentProcessRunning = false;
                                return;
                            }

                            //Get IE popup Title.
                            string popupName = dataset.Tables[0].Rows[i]["PopUpTitle"].ToString().Trim();
                            //Get IE popup Title.
                            string buttonName = dataset.Tables[0].Rows[i]["ButtonName"].ToString().Trim();
                            //Get the action need to be done.
                            string action = dataset.Tables[0].Rows[i]["Action"].ToString().Trim();

                            //Populate data if provided.
                            var data = dataset.Tables[0].Rows[i]["Data"].ToString().Trim();
                            if (windowTitleList.Contains(popupName))
                            {
                                //Switching control based on action given.
                                switch (action.ToLower())
                                {
                                    case "click":
                                        BtnClick(popupName, buttonName);
                                        break;
                                    case "keypress":
                                        KeyPress(popupName, data);
                                        break;
                                    case "close":
                                        _objAutoit.WinClose(popupName);
                                        break;
                                    case "run":
                                    case "execute":
                                        if (!File.Exists(_applicationPath + data))
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
                                            specialScriptProcess.StartInfo.WorkingDirectory = workingDirectory;
                                            specialScriptProcess.StartInfo.ErrorDialog = false;
                                            // Start the process
                                            specialScriptProcess.Start();
                                            specialScriptProcess.WaitForExit(10000);
                                        }
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
                Console.WriteLine(ConsoleMessages.REGISTER_AUTOIT);
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine(ConsoleMessages.INVALID_ARGUMNETS);
            }
            catch (Exception e)
            {
                Console.WriteLine(ConsoleMessages.PARALLEL_rECOVERY + e.Message + e.StackTrace);
            }

        }
        #endregion
        #region Private Static Methods.
        /// <summary>
        /// Key Press -> Send specified Key to window.
        /// </summary>
        /// <param name="popUpName">string : Name of popup to recover.</param>
        /// <param name="key"> string : Key to send.</param>
        private static void KeyPress(string popUpName, string key)
        {
            try
            {
                _objAutoit.WinActivate(popUpName, string.Empty); //Activate window.

                switch (key.ToLower())//Send Specified Key.
                {
                    case "tab":
                        _objAutoit.Send("{TAB}");
                        break;
                    case "space":
                        _objAutoit.Send("{SPACE}");
                        break;
                    case "enter":
                        _objAutoit.Send("{ENTER}");
                        break;
                    default:
                        _objAutoit.Send(key, 1);
                        break;
                }

            }
            catch (Exception e) 
            {
                KryptonException.Writeexception(e.InnerException);
            }
        }

        /// <summary>
        /// btnClick : Click on any object in specified window.
        /// </summary>
        /// <param name="popUpName">string : name of popup to recover.</param>
        /// <param name="btnTitle"></param>
        private static void BtnClick(string popUpName, string btnTitle = "")
        {
            try
            {
                _objAutoit.WinActivate(popUpName);

                //Click on control, close dialog if unsuccessful
                if (!(btnTitle.Equals("") || btnTitle.Equals(string.Empty)))
                {
                    if (_objAutoit.ControlClick(popUpName, "", "[TEXT:" + btnTitle + "]") == 0)
                    {
                    }
                }
            }
            catch (Exception e)
            {
                KryptonException.Writeexception(e.InnerException);
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
                if (fileExtension != null && fileExtension.ToLower().Equals(".csv"))
                {
                    if (!File.Exists(tmpFileName))
                    {
                        Console.WriteLine(ConsoleMessages.FOD + filePath);
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
                KryptonException.Writeexception(ex);
            }
            finally
            {
                File.Delete(tmpFileName);
            }
            return dataset;
        }


        private static DataSet GetExcelDataSet(string excelFilePath)
        {
            string tempFilePath = Path.GetTempFileName();
            IExcelDataReader excelReader = null;
            DataSet oResultSet = null;
            FileInfo oFileInfo = new FileInfo(excelFilePath);
            File.Copy(oFileInfo.FullName, tempFilePath, true);
            FileStream stream = File.OpenRead(tempFilePath);
            try
            {
                excelReader = oFileInfo.Extension.ToLower() == ".xls" ? ExcelReaderFactory.CreateBinaryReader(stream) : ExcelReaderFactory.CreateOpenXmlReader(stream);
                excelReader.IsFirstRowAsColumnNames = true;
                oResultSet = excelReader.AsDataSet();

                File.Delete(tempFilePath);
            }
            finally
            {
                if (excelReader != null)
                    excelReader.Close();
                stream.Close();
            }
            return oResultSet;
        }

        ///  <summary>
        /// Get Temp file specified.
        ///  </summary>
        ///  <param name="filePath" file extension>string : fileExtension.</param>
        ///  /// <param name="filePath" >string : Filepath.</param>
        /// <param name="fileExtension"></param>
        /// <param name="origPath"></param>
        /// <returns>Temp file.</returns>
        private static string GetTemporaryFile(string fileExtension, string origPath)
        {
            string response = string.Empty;
            try
            {
                if (!fileExtension.StartsWith("."))
                    fileExtension = "." + fileExtension;
                response = Path.GetTempPath() + Guid.NewGuid().ToString() + fileExtension;
                File.Copy(origPath, response);
            }
            catch (Exception)
            {
                // ignored
            }
            return response;
        }
        #endregion
    }
}
