/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: ExcelManager.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: Contain Method to Read Excel Sheet Test Cases.
*****************************************************************************/
using System;
using System.Data;
using System.Reflection;
using System.IO;
using System.Data.OleDb;
using ClosedXML.Excel;
using Excel;
using System.Collections.Generic;
using System.Text;

namespace ExcelLib
{   
  

    public class CSVGenericParser
    {
        public  DataSet ReadExcelData(string ExcelFilePath)
        {
            DataSet oResultDataSet = null;
            FileInfo oFileInfo = new FileInfo(ExcelFilePath);
            using (GenericParsing.GenericParserAdapter gp = new GenericParsing.GenericParserAdapter(ExcelFilePath, Encoding.UTF8))
            {
                gp.FirstRowHasHeader = true;
                gp.ColumnDelimiter = ',';
                oResultDataSet = gp.GetDataSet();
            }

            if (oResultDataSet != null && oResultDataSet.Tables.Count > 0)
                oResultDataSet.Tables[0].TableName = new FileInfo(ExcelFilePath).Name;

            return oResultDataSet;
        }
        
        public  void UpdateExcelData(string ExcelFilePath, string SheetName, DataTable SheetData, string FirstCellName, string LastColumnName)
        {
            var wb = new XLWorkbook(ExcelFilePath);
            int i = 0;
            foreach (var sheet in wb.Worksheets)
            {
                if (SheetName.ToLower().Equals(sheet.Name.ToLower()))
                    break;
                i++;
            }
            wb.Worksheets.Delete(i + 1);
            SheetData.TableName = SheetName;
            IXLWorksheet oIXLWorksheet = wb.Worksheets.Add(SheetData);
            oIXLWorksheet.Position = i + 1;
            wb.SaveAs(ExcelFilePath);
        }

        public  void AddNewSheetsInExcel(string ExcelFilePath, Dictionary<string, DataTable> SheetDataCollection, string FirstCellName, string LastColumnName)
        {
            var wb = new XLWorkbook(ExcelFilePath);
            foreach (string key in SheetDataCollection.Keys)
            {
                string keySheetName = key;
                SheetDataCollection[keySheetName].TableName = keySheetName;
                IXLWorksheet oIXLWorksheet = wb.Worksheets.Add(SheetDataCollection[key]);
            }
            wb.SaveAs(ExcelFilePath);
        }
    }

    internal class OledbExcelReader
    {
        private string ExcelConnection =
             "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;IMEX=1\"";

        public  DataSet ReadExcelData(string ExcelFilePath)
        {
            string tmpFileName = CopyToTempFile(ExcelFilePath);
            string connectionString = ExcelConnection.Replace("{0}", tmpFileName);

            DataSet strTestCaseFlow = GetRequiredRows(tmpFileName, connectionString);
            return strTestCaseFlow;
        }

        private  string CopyToTempFile(string oriFilePath)
        {
            string tmpFileName = string.Empty;
            FileInfo oFileInfo = new FileInfo(oriFilePath);
            tmpFileName = GetTemporaryFile(oFileInfo.Extension);
            File.Copy(oriFilePath, tmpFileName, true);
            return tmpFileName;
        }

        private  DataSet GetRequiredRows(string filePath, string connectionString)
        {
            DataSet dataset = new DataSet();
            //check whether given excel file exists or not before starting.           
            try
            {
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    foreach (DataRow dr in sheets.Rows)
                    {
                        int tempChek = dr["TABLE_NAME"].ToString().IndexOf("$") + 1;
                        if (tempChek < dr["TABLE_NAME"].ToString().Length)
                        {   if (!dr["TABLE_NAME"].ToString()[tempChek].Equals('\''))
                            {
                                if (!dr["TABLE_NAME"].ToString()[tempChek].Equals(string.Empty))
                                    continue;
                            }                            
                         }
                        using (var cmd = conn.CreateCommand())
                        {
                            DataTable dt = new DataTable();
                            dt.TableName = dr["TABLE_NAME"].ToString().Replace("$", string.Empty);
                            cmd.CommandText = "SELECT * FROM [" + dr["TABLE_NAME"].ToString() + "] ";

                            var adapter = new OleDbDataAdapter(cmd);
                            adapter.Fill(dt);
                            dataset.Tables.Add(dt);
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
            finally
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }

            return dataset;
        }

        private  string GetTemporaryFile(string extn)
        {
            string response = string.Empty;
            if (!extn.StartsWith("."))
                extn = "." + extn;

            response = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + extn;
            return response;
        }
    }

    internal class ExcelLibReader
    {
      public  DataSet ReadExcelData(string ExcelFilePath)
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
    }
}
