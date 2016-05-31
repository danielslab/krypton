/***************************************************************************
** Application: Krypton: ThinkSys Automation Generic System
** File Name: DBRelatedMethods.cs
** Copyright 2012, ThinkSys Software
** All rights reserved
** Developed by: ThinkSys Software
** Description: DB Method to Perform DB Operations.
*****************************************************************************/
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Threading;
using Common;

namespace TestDriver
{
    class DBRelatedMethods
    {
        private static OdbcCommand comm = null;
        private static string query = string.Empty;
        private static string ConnectionString = string.Empty;
        private static OdbcConnection conn = null;
        private static string errorMessage = string.Empty;
        /// <summary>
        // This method will fetch the data from a SQL database. It will be passed a query or a sql file
        /// that will contain the queries as a string. These queries will be performed on the database and the results
        /// will be saved to dictionary.
        /// </summary>
        /// <param name="selectQuery">It is a string that may either be a SQL query or the SQL file name.
        /// In case it is a query, it will be performed on the specified database.
        /// In case it is name of the SQL file, it will be read and queries inside this file will be performed on the database.</param>
        /// <returns>True on success else false</returns>
        public static bool GetDataFromDatabase(string selectQuery, string databaseType)
        {
            //Create a new connection string based on type of database specified in Parameters.ini file
            string connectionString = getDBConnectionString(databaseType);

            //Get location where SQL queries are stored
            string sqlQueryPath = Path.GetFullPath(Property.SqlQueryFilePath);

            // Check whether given argument is a query or sql query file.
                if (selectQuery.Split('.')[1].ToLower() == "sql")
                {
                    try
                    {
                        StreamReader readFile = new StreamReader(Path.Combine(sqlQueryPath, selectQuery));       
                        selectQuery = readFile.ReadToEnd();
                        readFile.Close();
                    }
                    catch (Exception exception)
                    {
                        Property.Remarks = Utility.GetCommonMsgVariable("KRYPTONERRCODE0026").Replace("{MSG}", exception.Message);
                        return false;
                    }
                }
            try
            {
                //replace the variables in the query by their values.
                selectQuery = Utility.ReplaceVariablesInString(selectQuery);
                query = selectQuery;
                ConnectionString = connectionString;

                OdbcDataReader reader = null;

                reader = executeQueryTillTimeout(); 

                bool isData = false;
                //now set new variables with fetched values from database.

                #region Store retrieved data into variables
                DataTable queryResults = new DataTable();
                queryResults.Load(reader);

                if (queryResults.Rows.Count > 0)
                {
                    isData = true;
                    int randomNumber = Utility.RandomNumber(1, queryResults.Rows.Count);
                    DataRow finalRecord = queryResults.Rows[randomNumber - 1];

                    for (int i = 0; i < queryResults.Columns.Count; i++)
                    {
                        Utility.SetVariable(queryResults.Columns[i].ColumnName, finalRecord[i].ToString());

                        //If values are being retrieved, show all retrived values in comments column
                        Property.Remarks = Property.Remarks +
                                                  queryResults.Columns[i].ColumnName + ":=" + finalRecord[i].ToString() + "; ";
                    }

                }


                //This section is replaced by section above in order to store possible unique values

                #endregion

                Property.Remarks = Property.Remarks + " Sql Query Used: " + selectQuery;
                comm.Dispose();
                reader.Close();
                conn.Close();

                if (isData == false)
                {
                    Property.Remarks = "No data could be retrieved from database, please check sql query" + "." + "\n" +
                                       "Connection String: " + connectionString + "\n" +
                                       "Sql Query Used: " + selectQuery;
                    return false;
                }
            }
            catch (TimeoutException)
            {
                Property.Remarks = Utility.GetCommonMsgVariable("KRYPTONERRCODE0058").Replace("{MSG}", errorMessage);
                return false;
            }
            catch (Exception exception)
            {
                Property.Remarks = Utility.GetCommonMsgVariable("KRYPTONERRCODE0027").Replace("{MSG}", exception.Message);
                Property.Remarks = Property.Remarks + "   Sql Query used : " + selectQuery;
                return false;
            }

            return true;
        }

        /// <summary>
        /// This method will update SQL database. It will be passed an update query or a sql file
        /// that will contain the update queries. These queries will require write access to the database.
        /// </summary>
        /// <param name="updateQuery">It is a string that may either be a SQL query or the SQL file name.
        /// In case it is a query, it will be performed on the specified database.
        /// In case it is name of the SQL file, it will be read and queries inside this file will be performed on the database.</param>
        /// <returns>True on success else false</returns>
        public static bool ExecuteDatabaseQuery(string updateQuery, string databaseType)
        {
            //Create a new connection string based on type of database specified in Parameters.ini file
            string connectionString = getDBConnectionString(databaseType);

            //Get location where SQL queries are stored
            string sqlQueryPath = Path.GetFullPath(Property.SqlQueryFilePath);               

            int rowsAffected = 0;

            // Check whether given argument is a query or sql query file.
            if (updateQuery.Split('.')[1].ToLower() == "sql")
            {
                try
                {
                    StreamReader readFile = new StreamReader(sqlQueryPath + "/" + updateQuery);
                    updateQuery = readFile.ReadToEnd();
                    readFile.Close();
                }
                catch (Exception exception)
                {
                    Property.Remarks = Utility.GetCommonMsgVariable("KRYPTONERRCODE0026").Replace("{MSG}", exception.Message);
                    return false;
                }
            }

            try
            {
                //replace the variables in the query by their values.
                updateQuery = Utility.ReplaceVariablesInString(updateQuery);

                //using odbc connection to work with multiple database types
                OdbcConnection conn = new OdbcConnection(connectionString);
               
                conn.Open();

                //execute the update query on the database.                
                comm = new OdbcCommand(updateQuery, conn);
                rowsAffected = comm.ExecuteNonQuery();

                Property.Remarks = " Sql Query Used: " + updateQuery;

                comm.Dispose();
                conn.Close();
            }
            catch (Exception exception)
            {
                bool updateResult = ExecuteUpdateQueryTillTimeout(connectionString, updateQuery);
                if (!updateResult)
                {
                    Property.Remarks = Utility.GetCommonMsgVariable("KRYPTONERRCODE0028").Replace("{MSG}", exception.Message);
                    Property.Remarks = Property.Remarks + "   SQL Query used : " + updateQuery;
                    return false;
                }
            }

            return true;
        }

        private static bool updateCondition(string connectionString, string updateQuery)
        {
            try
            {
                int rowsAffected = 0;
                OdbcConnection conn = new OdbcConnection(connectionString);
                conn.Open();
                comm = new OdbcCommand(updateQuery, conn);
                rowsAffected = comm.ExecuteNonQuery();
                comm.Dispose();
                conn.Close();
                return true;
            }
            catch { return false; }
        }
        /// <summary>
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="updateQuery"></param>
        /// <returns>bool value</returns>
        private static bool ExecuteUpdateQueryTillTimeout(string connectionString, string updateQuery)
        {
            {
                double diff = 0;
                double start_time = 0;
                double end_time = 0;
                int maxTime = int.Parse(Utility.GetParameter("GlobalTimeout"));
                try
                {
                    for (int seconds = 0; ; seconds++)
                    {
                        start_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                        if (seconds >= maxTime || diff >= maxTime)
                        {
                            throw new TimeoutException();
                        }
                        bool readerResult = updateCondition(connectionString, updateQuery);
                        if (readerResult)
                        {
                            return readerResult;
                        }
                        Thread.Sleep(1000);
                        end_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                        diff = diff + (end_time - start_time);
                    }
                }
                catch (TimeoutException)
                {
                    return false;
                }
            }
        }

        /// <summary>
        ///  Verify database method is intended to validate if correct values are stored in backend systems.
        /// </summary>
        /// <param name="data">This will be record number for which points to sql query stored in test case/ test data worksheet.
        /// E.g. which data column mentiods 2, it means locate a worksheet named "DBTestData", find row, in which first column contains value 2,
        /// in that row, locate column 'SqlQuery' and perform actions
        /// </param>
        /// <param name="databaseType">This is a hint telling which database server is to be used</param>
        /// <returns></returns>
        public static bool VerifyDatabase(string data, string databaseType)
        {
            bool result = false;

            //Retrieve Sql query to be executed from test data
            string sqlQuery = Utility.GetVariable("[td]SQLQuery");

            //Create a new connection string based on type of database specified in Parameters.ini file
            string connectionString = getDBConnectionString(databaseType);

            try
            {
                string param = "[td]param";
                string expectedValue = string.Empty;
                string actualValue = string.Empty;

                //replace the variables in the query by their values.
                sqlQuery = Utility.ReplaceVariablesInString(sqlQuery);

                
                query = sqlQuery;
                ConnectionString = connectionString;

                //using odbc connection to work with multiple database types
                
                OdbcDataReader reader = null;

                //execute the query on the database.
                reader = executeQueryTillTimeout();

                
                //Handling when no data was retrieved from database
                if (!reader.HasRows)
                {
                    Property.Remarks = "No data could be retrieved from database, please check sql query" + "." + "\n" +
                                       "Connection String: " + connectionString + "\n" +
                                       "Sql Query Used: " + sqlQuery;
                    comm.Dispose();
                    reader.Close();
                    conn.Close();
                    return false;
                }

                //Compare output of each column with expected values from parameters
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        //Retrieve expected value from parameters
                        expectedValue = Utility.GetVariable(param + (i + 1).ToString());
                        expectedValue = Utility.ReplaceVariablesInString(expectedValue);

                        //Retrieve actual value from output of query
                        actualValue = reader[i].ToString();

                        //Compare both expected and actual values, these should equal
                        if (Utility.DoKeywordMatch(expectedValue, actualValue))
                        {
                            result = true;
                            Property.Remarks = "Actual value " + actualValue + " of column '" + reader.GetName(i) +
                                                      "' matches with expected value " + expectedValue + "\n" +
                                            "Connection String: " + connectionString + "\n" +
                                            "Sql Query Used: " + sqlQuery;
                        }
                        else
                        {
                            result = false;
                            Property.Remarks = "Actual value " + actualValue + " of column '" + reader.GetName(i) +
                                                        "' does not matches with expected value " + expectedValue + "\n" +
                                              "Connection String: " + connectionString + "\n" +
                                              "Sql Query Used: " + sqlQuery;
                            break;
                        }
                    }

                    //Break loop as only first record needs to be validated
                    break;
                }

                comm.Dispose();
                reader.Close();
                conn.Close();
                return result;
            }
            catch (TimeoutException)
            {
                Property.Remarks = Utility.GetCommonMsgVariable("KRYPTONERRCODE0058").Replace("{MSG}", errorMessage) + "\n" + "Connection String: " + connectionString + "\n" +
                                   "Sql Query Used: " + sqlQuery;
               
                return false;
            }
            catch (Exception e)
            {
                //throw e;
                //Custom messaging to be returned
                result = false;
                Property.Remarks = "Exception in database verification: " + e.Message + "\n" +
                                   "Connection String: " + connectionString + "\n" +
                                   "Sql Query Used: " + sqlQuery;
                return result;
            }

        }

        /// <summary>
        /// This method will create a new sql server connection string based on db hint passed on.
        /// </summary>
        /// <param name="dbHint">This is a pointer that tell which database server to use. 
        /// Actual value of database server will be available from test environment specific parameter file.
        /// </param>
        /// <returns>Finally usable connection string </returns>
        public static string getDBConnectionString(string dbHint)
        {
            try
            {
                //check if a dbHint is actually passed, if not, fall back to default one
                if (dbHint.Equals(string.Empty))
                {
                    dbHint = Utility.GetParameter("DefaultDb");
                }

                //Retrieve actual name of database server from variables
                string dbServerName = Utility.GetVariable(dbHint);

                //Retrieve connection string signature from Parameters.ini file
                string dbConnectionString = Utility.GetParameter("DBConnectionString");

                //Replace variables in connection string
                dbConnectionString = dbConnectionString.Replace("{DbServerName}", dbServerName);
                dbConnectionString = Utility.ReplaceVariablesInString(dbConnectionString);

                //Return this connection string created so far
                return dbConnectionString;

            }
            catch (Exception e)
            {
                throw e;
            }

        }

        /// <summary>
        /// </summary>
        /// <returns>OdbcDataReader : reader obtained. </returns>
        private static OdbcDataReader executeQueryTillTimeout()
        {
            {
                double diff = 0;
                double start_time = 0;
                double end_time = 0;
                int maxTime = int.Parse(Utility.GetParameter("GlobalTimeout"));
                try
                {
                    for (int seconds = 0; ; seconds++)
                    {
                        start_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                        if (seconds >= maxTime || diff >= maxTime)
                        {
                            throw new TimeoutException();
                        }
                        OdbcDataReader readerResult = conditionalAttachedMethod();
                        if (readerResult != null)
                        {
                            return readerResult;
                        }
                        Thread.Sleep(1000);
                        end_time = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                        diff = diff + (end_time - start_time);
                    }
                }
                catch (TimeoutException e)
                {
                    conn.Close();
                    throw e;
                }
            }
        }

        private static OdbcDataReader conditionalAttachedMethod()
        {
            try
            {
                conn = new OdbcConnection(ConnectionString);
                conn.Open();
                comm = new OdbcCommand(query, conn);                
                comm.CommandTimeout = int.Parse(Utility.GetParameter("GlobalTimeout"));
                OdbcDataReader dr= comm.ExecuteReader();
                if (dr != null && !dr.HasRows)
                {
                    dr.Close();
                    conn.Close();
                    Thread.Sleep(2500);
                    conn = new OdbcConnection(ConnectionString);
                    conn.Open();
                    comm = new OdbcCommand(query, conn);
                    comm.CommandTimeout = int.Parse(Utility.GetParameter("GlobalTimeout"));
                    dr = comm.ExecuteReader();
                }
                return dr;
            }
            catch (Exception e)
            {
                conn.Close();
                errorMessage = string.Empty;
                errorMessage = e.Message;
                return null;
            }
        }
    }
}
