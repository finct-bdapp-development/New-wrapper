using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
namespace New_Wrapper
{

    /// <summary>
    /// Handles the connecting to and updating of data using OLE
    /// </summary>
    /// //<remarks>The default instance of the data handler only connects and updates a single table. For datasets containing
    /// multiple updateable tables use the MultiHandler</remarks>
    public class DataHandler
    {

        #region Properties

        SqlCommand _ActionedCommand;
        /// <summary>
        /// An SQL command onbject containing the result of a data connection
        /// </summary>
        public SqlCommand ActionedCommand
        {
            get { return _ActionedCommand; }
            set { _ActionedCommand = value; }
        }

        #endregion

        #region Genric Excel

        /// <summary>
        /// Returns the data in an Excel sheet as a data table
        /// </summary>
        /// <param name="pathToExcelFile">The path to the file to be converted</param>
        /// <param name="sheetName">The name of the sheet where the data is held</param>
        /// <param name="includesHeaderRow">Whether the data has a header row showing the column names</param>
        /// <returns></returns>
        /// <remarks>The method uses the ACE provider and should work for all Excel files</remarks>
        public DataTable ReturnExcelSheetAsDataTable(string pathToExcelFile, string sheetName, Boolean includesHeaderRow)
        {
            string header = "No";
            if (includesHeaderRow == true) header = "Yes";
            string connection = "Provider=Microsoft.JET.OLEDB.4.0;Data Source =" + pathToExcelFile + ";Extended Properties = \"Excel 8.0;HDR =" + header + ";IMEX=1\"";
            string query = "SELECT * FROM [" + sheetName + "$]";
            New_Wrapper.DataHandler myHandler = new New_Wrapper.DataHandler(connection, query);
            return myHandler.CreateDataset().Tables[0];
        }

        #endregion

        #region SQL stored procedures

        /// <summary>
        /// Action a stored procedure with associated parameters
        /// </summary>
        /// <param name="parameters">A list of the parameters to apply</param>
        /// <param name="connectionString">the connection string for the SQL data source</param>
        /// <param name="storedProcedureName">The name of the stored procedure</param>
        public void RunStoredProcedure(ref List<SqlParameter> parameters, string connectionString, string storedProcedureName)
        {
            SqlConnection pvConnection;
            try
            {
                pvConnection = new SqlConnection(connectionString);
            }
            catch
            {
                throw;
            }
            //SqlCommand pvCommand = new SqlCommand(storedProcedureName, pvConnection);
            //SqlCommand 
            try
            {
                ActionedCommand = new SqlCommand(storedProcedureName, pvConnection);

                ActionedCommand.CommandType = CommandType.StoredProcedure;

                ActionedCommand.Parameters.Clear();
                foreach (SqlParameter item in parameters)
                {
                    ActionedCommand.Parameters.Add(item);
                }

                pvConnection.Open();
                ActionedCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            //return ActionedCommand;
        }

        /// <summary>
        /// Action a stored procedure that does not require any parameters
        /// </summary>
        /// <param name="connectionString">The connection string for the SQL data source</param>
        /// <param name="storedProcedureName">The name of the stored procedure</param>
        public void RunStoredProcedure(string connectionString, string storedProcedureName)
        {
            SqlConnection pvConnection;
            try
            {
                pvConnection = new SqlConnection(connectionString);
            }
            catch
            {
                throw;
            }
            //SqlCommand pvCommand = new SqlCommand(storedProcedureName, pvConnection);
            //SqlCommand 
            ActionedCommand = new SqlCommand(storedProcedureName, pvConnection);

            ActionedCommand.CommandType = CommandType.StoredProcedure;

            pvConnection.Open();
            ActionedCommand.ExecuteNonQuery();
        }

        /// <summary>
        /// Confirm whether the currently logged in user has the specified role
        /// </summary>
        /// <param name="provider">The connection string for the data source</param>
        /// <param name="roleName">The name of the SQL user role to be checked</param>
        /// <returns></returns>
        public Boolean CheckUserHasRole(string provider, string roleName)
        {
            New_Wrapper.DataHandler myHandler = new New_Wrapper.DataHandler();
            List<SqlParameter> paramArray = new List<SqlParameter>();
            SqlParameter myParam = new SqlParameter("@RoleName", SqlDbType.NVarChar, 50);
            myParam.Value = roleName;
            paramArray.Add(myParam);
            myHandler.RunStoredProcedure(ref paramArray, provider, "ReturnRoleMemberStatusCurrentUser");
            DataTable temp = new DataTable();
            temp.Load(myHandler.ActionedCommand.ExecuteReader());
            if (temp.Rows[0]["HasRole"].ToString() == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Return a dataset based on the result of running a stored procedure that returns multiple results
        /// </summary>
        /// <param name="provider">The connection string for the data source</param>
        /// <param name="parameters">The parameters required by the stored procedure</param>
        /// <param name="procedureName">The anme of the stored procedure</param>
        /// <returns></returns>
        public DataSet RunStoredProcedurehAndReturnDataset(string provider, ref List<SqlParameter> parameters, string procedureName)
        {
            RunStoredProcedure(ref parameters, provider, procedureName);
            DataSet temp = new DataSet();
            System.Data.SqlClient.SqlDataAdapter adp = new SqlDataAdapter();
            adp.SelectCommand = ActionedCommand;
            adp.Fill(temp);
            return temp;
        }

        /// <summary>
        /// Returns a data table based on running a stored procedure that has a single result set
        /// </summary>
        /// <param name="provider">The connection string for the data source</param>
        /// <param name="procedureName">The name of the stored procedure</param>
        /// <returns></returns>
        public DataTable RunStoredProcedureAndReturnSingleTable(string provider, string procedureName)
        {
            RunStoredProcedure(provider, procedureName);
            DataTable tempData = new DataTable();
            tempData.Load(ActionedCommand.ExecuteReader(CommandBehavior.CloseConnection));
            return tempData;
        }

        /// <summary>
        /// Returns a data table based on running a stored procedure that has a single result set
        /// </summary>
        /// <param name="provider">The connection string for the data source</param>
        /// <param name="parameters">The parameters required by the stored procedure</param>
        /// <param name="procedureName">The name of the stored procedure</param>
        /// <returns></returns>
        public DataTable RunStoredProcedureAndReturnSingleTable(string provider, ref List<SqlParameter> parameters, string procedureName)
        {
            RunStoredProcedure(ref parameters, provider, procedureName);
            DataTable tempData = new DataTable();
            tempData.Load(ActionedCommand.ExecuteReader(CommandBehavior.CloseConnection));
            return tempData;
        }

        /// <summary>
        /// Returns a data table based on running a stored procedure that has multiple results
        /// </summary>
        /// <param name="provider">The connection string for the data source</param>
        /// <param name="parameters">The parameters required by the stored procedure</param>
        /// <param name="procedureName">The name of the stored procedure</param>
        /// <param name="tableToReturn">The result set that should be returned</param>
        /// <returns></returns>
        /// <remarks>This method can result in large amounts of unneccessary data being returned. Where possible use the overload that uses a
        /// stored procedure that returns a single result set</remarks>
        public DataTable RunStoredProcedureAndReturnSingleTable(string provider, ref List<SqlParameter> parameters, string procedureName, int tableToReturn)
        {
            return RunStoredProcedurehAndReturnDataset(provider, ref parameters, procedureName).Tables[tableToReturn];
        }

        /// <summary>
        /// Confirm whether a stroed procedure already exists in the specified parameter array
        /// </summary>
        /// <param name="parameters">A list array of SQL parameters</param>
        /// <param name="parameterToAdd">The sql parameter to be checked against the existing parameters</param>
        /// <returns></returns>
        /// <remarks>The comparison is based on the parameter name and value. If the name exists but the value is different an
        /// exception will be thrown</remarks>
        public Boolean IsExistingParameter(ref List<SqlParameter> parameters, ref SqlParameter parameterToAdd)
        {
            foreach (SqlParameter item in parameters)
            {
                if (item.ParameterName == parameterToAdd.ParameterName)
                {
                    if (item.Value != parameterToAdd.Value)
                    {
                        throw new Exception("An attempt has been made to create a parameter (" + parameterToAdd.ParameterName + ") which already exists but has a different value.");
                    }
                    return true;
                }
            }
            return false;
        }

        #endregion

        #region SQL Bulk Inserts

        /// <summary>
        /// Import bulk data into SQL database table
        /// </summary>
        /// <param name="passedDataSourceProvider">The OLEDb dataprovider to the source data to be inserted in to SQL</param>
        /// <param name="passedDataSourceQueryString">If query not included in provider string, include here</param>
        /// <param name="passedDestinationProvider">The SQL dataprovider to the source data to be inserted in to SQL</param>
        /// <param name="passedDestinationTableName">The name of the destination table name in the SQL Database</param>
        /// <param name="passedNumberOfFieldsInLists">This nuumber must match the number of items in passed source and desination lists</param>
        /// <param name="passedSourceFieldNames">The names of the field in the source data that will be inserted in to destination table</param>
        /// <param name="passedDestinationFieldNames">The corresponding field names in the destination table that will receive the source data</param>
        /// <param name="headerRowInSourceProvider">False if header row not included. This will ensure delete first row of data table is deleted.</param>
        /// <param name="passedBulkCopyTimeout">Default server value = 30. Consider increasing for large datasets</param>
        public void InsertBulkData(string passedDataSourceProvider, string passedDataSourceQueryString, string passedDestinationProvider, string passedDestinationTableName, int passedNumberOfFieldsInLists, ref List<string> passedSourceFieldNames, ref List<string> passedDestinationFieldNames, Boolean headerRowInSourceProvider, int passedBulkCopyTimeout)
        {
            try
            {
                using (OleDbConnection dataSource_con = new OleDbConnection(passedDataSourceProvider))
                {
                    dataSource_con.Open();
                    DataTable dtSourceData = new DataTable();
                    using (OleDbDataAdapter oda = new OleDbDataAdapter(@passedDataSourceQueryString, dataSource_con))
                    {
                        oda.Fill(dtSourceData);
                        if (headerRowInSourceProvider == false)
                        {
                            dtSourceData.Rows[0].Delete();
                        }
                    }
                    dataSource_con.Close();

                    string consString = passedDestinationProvider;
                    using (SqlConnection con = new SqlConnection(consString))
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            sqlBulkCopy.BulkCopyTimeout = passedBulkCopyTimeout;
                            sqlBulkCopy.DestinationTableName = passedDestinationTableName;
                            for (int i = 0; i < passedNumberOfFieldsInLists; i++)//20/07/2016 removed the "-1" from int i = 0; i < passedNumberOfFieldsInLists - 1; i++
                            {
                                sqlBulkCopy.ColumnMappings.Add(passedSourceFieldNames[i].ToString(), passedDestinationFieldNames[i].ToString());
                            }
                            con.Open();
                            sqlBulkCopy.WriteToServer(dtSourceData);
                            con.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred in the bulk data import:\r\nError Message - " + ex.Message + "\r\nStack Trace - " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Import bulk data into SQL database table
        /// </summary>
        /// <param name="passedDataTable">Datatable containing the data to insert into SQL database</param>
        /// <param name="passedDestinationProvider">The SQL dataprovider to the source data to be inserted in to SQL</param>
        /// <param name="passedDestinationTableName">The name of the destination table name in the SQL Database</param>
        /// <param name="passedNumberOfFieldsInLists">This nuumber must match the number of items in passed source and desination lists</param>
        /// <param name="passedSourceFieldNames">The names of the field in the source data that will be inserted in to destination table</param>
        /// <param name="passedDestinationFieldNames">The corresponding field names in the destination table that will receive the source data</param>
        /// <param name="headerRowInSourceProvider">False if header row not included. This will ensure delete first row of data table is deleted.</param>
        /// <param name="passedBulkCopyTimeout">Default server value = 30. Consider increasing for large datasets</param>
        public void InsertBulkData(DataTable passedDataTable, string passedDestinationProvider, string passedDestinationTableName, int passedNumberOfFieldsInLists, ref List<string> passedSourceFieldNames, ref List<string> passedDestinationFieldNames, Boolean headerRowInSourceProvider, int passedBulkCopyTimeout)
        {
            try
            {
                DataTable dtSourceData = passedDataTable;
                string consString = passedDestinationProvider;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        sqlBulkCopy.BulkCopyTimeout = passedBulkCopyTimeout;
                        sqlBulkCopy.DestinationTableName = passedDestinationTableName;
                        for (int i = 0; i < passedNumberOfFieldsInLists; i++)//20/07/2016 removed the "-1" from int i = 0; i < passedNumberOfFieldsInLists - 1; i++
                        {
                            sqlBulkCopy.ColumnMappings.Add(passedSourceFieldNames[i].ToString(), passedDestinationFieldNames[i].ToString());
                        }
                        con.Open();
                        sqlBulkCopy.WriteToServer(dtSourceData);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred in the bulk data import:\r\nError Message - " + ex.Message + "\r\nStack Trace - " + ex.StackTrace);
            }
        }
        /// <summary>
        /// As above but as a transaction to rollback
        /// </summary>
        /// <param name="passedDataTable"></param>
        /// <param name="passedDestinationProvider"></param>
        /// <param name="passedDestinationTableName"></param>
        /// <param name="passedNumberOfFieldsInLists"></param>
        /// <param name="passedSourceFieldNames"></param>
        /// <param name="passedDestinationFieldNames"></param>
        /// <param name="headerRowInSourceProvider"></param>
        /// <param name="passedBulkCopyTimeout"></param>
        public void InsertBulkDataBatch(DataTable passedDataTable, string passedDestinationProvider, string passedDestinationTableName, int passedNumberOfFieldsInLists, ref List<string> passedSourceFieldNames, ref List<string> passedDestinationFieldNames, Boolean headerRowInSourceProvider, int passedBulkCopyTimeout, int batchSize)
        {
            try
            {
                DataTable dtSourceData = passedDataTable;
                string consString = passedDestinationProvider;
                using (SqlConnection con = new SqlConnection(consString))
                {
                    con.Open();
                    using (SqlTransaction transaction = con.BeginTransaction())
                    {
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con,SqlBulkCopyOptions.Default,transaction))
                        {
                            sqlBulkCopy.BatchSize = batchSize;
                            sqlBulkCopy.BulkCopyTimeout = passedBulkCopyTimeout;
                            sqlBulkCopy.DestinationTableName = passedDestinationTableName;
                            for (int i = 0; i < passedNumberOfFieldsInLists; i++)//20/07/2016 removed the "-1" from int i = 0; i < passedNumberOfFieldsInLists - 1; i++
                            {
                                sqlBulkCopy.ColumnMappings.Add(passedSourceFieldNames[i].ToString(), passedDestinationFieldNames[i].ToString());
                            }
                           
                            try
                            {
                                sqlBulkCopy.WriteToServer(dtSourceData);
                                transaction.Commit();
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                                transaction.Rollback();
                            }
                            finally
                            {

                                con.Close();
                            }
                          
                        }
                    }
                }


               
                   
                
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred in the bulk data import:\r\nError Message - " + ex.Message + "\r\nStack Trace - " + ex.StackTrace);
            }
        }
        /// <summary>
        /// Import bulk data into SQL database table
        /// </summary>
        /// <param name="passedData">The data to be imported</param>
        /// <param name="destinationProvider">The connection string for the datasource that the data will be uploaded to</param>
        /// <param name="destinationTable">The name of the table that data will be uploaded to</param>
        /// <remarks>The passed data must only contain the fields which are to be uploaded to the datasource</remarks>
        public void InsertBulkData(ref DataTable passedData, string destinationProvider, string destinationTable)
        {
            List<string> columnNames = new List<string>();
            foreach(DataColumn col in passedData.Columns)
            {
                columnNames.Add(col.ColumnName.ToString());
            }
            InsertBulkData(passedData, destinationProvider, destinationTable, columnNames.Count, ref columnNames, ref columnNames, false, 30);
        }

        /// <summary>
        ///  Send multiple parameter lists for batch bulk copy
        /// </summary>
        /// <param name="passedDataTableList">List of all tables</param>
        /// <param name="passedDestinationProvider">Provider</param>
        /// <param name="passedDestinationTableNameList">List of destination tables</param>
        /// <param name="headerRowInSourceProviderList">List of header row exists</param>
        /// <param name="passedSourceFieldNamesList">List of array of passed field names</param>
        /// <param name="passedDestinationFieldNamesList">List of array of destination field names</param>
        /// <param name="passedBulkCopyTimeout">Bulk copy timeout</param>
        /// <param name="passedBatchSizeList">List of batch sizes</param>
        public void InsertBulkDataBatch(string passedDestinationProvider, 
            DataTable[] passedDataTableList, 
            string[] passedDestinationTableNameList, 
            Boolean[] headerRowInSourceProviderList, 
            List<string[]> passedSourceFieldNamesList,
            List<string[]> passedDestinationFieldNamesList,
            int[] passedBatchSizeList,
            int passedBulkCopyTimeout=-1)
        {
            int tableCount = passedDataTableList.Count();
            int tableNameCount = passedDestinationTableNameList.Count();
            int headerRowCount = headerRowInSourceProviderList.Count();
            int sourceFieldsCount = passedSourceFieldNamesList.Count();
            int destFieldCount = passedDestinationFieldNamesList.Count();
            int batchSizeCount = passedBatchSizeList.Count();

            //Checks if passed lists contain same number of elements, if not - FAIL
            if (tableCount == tableNameCount && tableCount == headerRowCount && tableCount == sourceFieldsCount &&
               tableCount == destFieldCount && tableCount == batchSizeCount)
            {
                //Check if source and destination item lists are same size, if not - FAIL
                if (passedSourceFieldNamesList == passedDestinationFieldNamesList)
                {
                    try
                    {
                        //Iterate through each table
                        for (int y = 0; y < passedDataTableList.Count(); y++)
                        {
                            DataTable dtSourceData = passedDataTableList[y];
                            string consString = passedDestinationProvider;
                            using (SqlConnection con = new SqlConnection(consString))
                            {
                                con.Open();
                                using (SqlTransaction transaction = con.BeginTransaction())
                                {
                                    //Start copy using index of data table to access same index in other lists
                                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.Default, transaction))
                                    {
                                        if (passedBulkCopyTimeout > -1)
                                        {
                                            sqlBulkCopy.BulkCopyTimeout = passedBulkCopyTimeout;
                                        }

                                        sqlBulkCopy.BatchSize = passedBatchSizeList[y];
                                        sqlBulkCopy.DestinationTableName = passedDestinationTableNameList[y];

                                        //Get string array from List<string[]>
                                        string[] passedSourceFieldNames = passedSourceFieldNamesList[y];
                                        string[] passedDestinationFieldNames = passedDestinationFieldNamesList[y];

                                        for (int i = 0; i < passedSourceFieldNames.Count(); i++)//20/07/2016 removed the "-1" from int i = 0; i < passedNumberOfFieldsInLists - 1; i++
                                        {
                                            sqlBulkCopy.ColumnMappings.Add(passedSourceFieldNames[i].ToString(),
                                                passedDestinationFieldNames[i].ToString());
                                        }
                                        try
                                        {
                                            sqlBulkCopy.WriteToServer(dtSourceData);
                                            transaction.Commit();
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                            transaction.Rollback();
                                            con.Close();
                                            // If one bulk fails - stop further bulks
                                            throw new Exception("An error occurred in the bulk data import:\r\nError Message - " + 
                                                ex.Message + "\r\nStack Trace - " + ex.StackTrace);
                                        }
                                        finally
                                        {
                                            con.Close();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("An error occurred in the bulk data import:\r\nError Message - " + ex.Message + "\r\nStack Trace - " + ex.StackTrace);
                    }
                }
                else
                {
                    throw new Exception("An error occurred in the bulk data import:\r\nDestination and source field counts do not match");
                }
            }
            else
            {
                throw new Exception("An error occurred in the bulk data import:\r\nMultiple lists counts do not match");

            }
        }
        /// <summary>
        /// Import bulk data from a CSV file
        /// </summary>
        /// <param name="targetProvider">The connection string for the SQL data source into which the data will be inserted</param>
        /// <param name="targetTable">The name of the table that data will be inserted into</param>
        /// <param name="fileToImport">The full path to csv file to be imported</param>
        /// <param name="settingsFile">The full path to the settings file for mapping the CSV file to the data source</param>
        /// <param name="includesHeader">Whether the CSV file includes a header row</param>
        /// <remarks>The settings file is a simple ini file consisting of a delimited line showing where each field in the CSV
        /// should go within the datasource - CSVColumnName,AssociatedSQLColumnName
        /// For example:
        /// TPName,CustomerName
        /// Address1,AddressLine1</remarks>
        public void InsertBulkDataViaCSV(string targetProvider, string targetTable, string fileToImport, string settingsFile, Boolean includesHeader)
        {
            if(System.IO.File.Exists(fileToImport) == false)
            {
                throw new Exception("The file that you are trying to import either does not exist or you do not have permission to access it.");
            }
            if(System.IO.File.Exists(settingsFile)==false)
            {
                throw new Exception("The settings file for the import either does not exist or you do not have permission to access it.");
            }
            //Code to create the arrays, etc
            List<string> sourceFields = new List<string>();
            List<string> destinationFields = new List<string>();
            //Read setings file
            System.IO.StreamReader myReader = new System.IO.StreamReader(settingsFile);
            while(myReader.EndOfStream == false)
            {
                string[] lineText = myReader.ReadLine().Split(',');
                sourceFields.Add(lineText[0]);
                destinationFields.Add(lineText[1]);
            }
            myReader.Close();

            string headerExists = string.Empty;
            if(includesHeader == true)
            {
                headerExists = "YES";
            }
            else
            {
                headerExists = "NO";
            }
            string provider = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + System.IO.Path.GetDirectoryName(fileToImport) + @"\; Extended Properties = '" + "text;HDR=" + headerExists + ";FMT=Delimited';";
            DataHandler myHandler = new DataHandler(provider, "SELECT * FROM [" + System.IO.Path.GetFileName(fileToImport) + "]");
            DataTable temp = myHandler.CreateDataset().Tables[0];
            //HACK For some reason schema is resulting in the header row being added and a blank row with the header
            //Remove any rows where this happens
            for (int counter = 0; counter < temp.Rows.Count; counter++)
            {
                if (temp.Rows[counter][0].ToString() == temp.Columns[0].ColumnName)
                {
                    temp.Rows.RemoveAt(counter);
                    counter--;
                }
            }
            myHandler = null;
            InsertBulkData(temp, targetProvider, targetTable, sourceFields.Count(), ref sourceFields, ref destinationFields, false, 30);
            temp = null;
        }

        #endregion

        #region Convert CSV to data table
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileToImport"></param>
        /// <param name="includesHeader"></param>
        /// <returns></returns>
        public DataTable ReturnCSVAsDataTable(string fileToImport, Boolean includesHeader)
        {
            if (System.IO.File.Exists(fileToImport) == false)
            {
                throw new Exception("The file that you are trying to import either does not exist or you do not have permission to access it.");
            }

            string headerExists = string.Empty;
            if (includesHeader == true)
            {
                headerExists = "YES";
            }
            else
            {
                headerExists = "NO";
            }
            string provider = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + System.IO.Path.GetDirectoryName(fileToImport) + @"\; Extended Properties = '" + "text;HDR=" + headerExists + ";FMT=Delimited';";
            DataHandler myHandler = new DataHandler(provider, "SELECT * FROM [" + System.IO.Path.GetFileName(fileToImport) + "]");
            DataTable temp = myHandler.CreateDataset().Tables[0];
            return temp;
        }

        #endregion


        /// <summary>
        /// 
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public DataTable ConvertExcelToDataTable(string FileName)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }

        #region Convert Excel sheet to data table
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileToImport"></param>
        /// <returns></returns>
        public DataTable ReturnExcelAsDataTable(string fileToImport)
        {
            DataTable dt = new DataTable();

           /* dt=ConvertExcelToDataTable(fileToImport);
            return dt;*/
            string sampleProvider = string.Empty;
            string selectQuery = string.Empty;
            switch (System.IO.Path.GetExtension(fileToImport))
            {
                case ".xls":
                    sampleProvider = @"Provider = Microsoft.jet.oledb.4.0; Data Source = " + fileToImport + "; Extended Properties = 'Excel 8.0; HDR = Yes; IMEX = 1'"; //+ Properties.Settings.Default.RootPath + @"\Book1.xls; Extended Properties = 'Excel 8.0; HDR = Yes; IMEX = 1'";
                    selectQuery = "SELECT * FROM [SHEET1$]";
                    break;
                case ".xlsx":
                    sampleProvider = @"Provider = Microsoft.ACE.oledb.12.0; Data Source = " + fileToImport + "; Extended Properties = 'Excel 12.0 Xml; HDR = Yes'"; //+ Properties.Settings.Default.RootPath + @"\Book1.xls; Extended Properties = 'Excel 8.0; HDR = Yes; IMEX = 1'";
                    selectQuery = "SELECT * FROM [SHEET1$]";
                    break;
                case ".csv":
                    sampleProvider = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + System.IO.Path.GetDirectoryName(fileToImport) + "; Extended Properties ='text; HDR = Yes; FMT = Delimited'";
                    selectQuery = "SELECT * FROM [" + System.IO.Path.GetFileName(fileToImport) + "]";
                    break;
                case ".xml":
                    break;
                default:
                    return dt;
            }
            try
            {
                OleDbConnection myConnection;
                OleDbDataAdapter myCommand;
                myConnection = new OleDbConnection(sampleProvider);
                myCommand = new OleDbDataAdapter(selectQuery, myConnection);
                myCommand.TableMappings.Add("Table", "MainTable");
                try
                {
                    myCommand.Fill(dt);
                }
                catch (Exception ex)
                {
                    
                }
                myConnection.Close();
                return dt;

            }
            catch (Exception ex)
            {
                return dt;
            }

        }
        #endregion

        #region Write datatable to Excel



        #endregion

        #region Write datatable to CSV

        /// <summary>
        /// Dumps the contents of a data table to a comma separated file
        /// </summary>
        /// <param name="data">The table containing the data</param>
        /// <param name="path">The destination for the csv file</param>
        /// <param name="fileName">The name of the csv that will be produced</param>
        /// <remarks>This version assumes that the data does not contain field data containing commas as this will throw out the formatting of the row
        /// within the csv file. If necessary an overload of the method should be produced to accomodate data that contains commas and control 
        /// characters</remarks>
        public void GenerateCSV(DataTable data, string path, string fileName)
        {
            System.IO.StreamWriter myWriter = new System.IO.StreamWriter(path + @"\" + fileName);
            StringBuilder row = new StringBuilder();
            foreach(DataColumn item in data.Columns)
            {
                if (string.IsNullOrEmpty(row.ToString()))
                {
                    row.Append(item);
                }
                else
                {
                    row.Append("," + item.ColumnName);
                }
            }
            myWriter.WriteLine(row);
            foreach(DataRow item in data.Rows)
            {
                List<string> rowData = new List<string>();
                for (int col = 0; col < item.ItemArray.Count(); col++)
                {
                    rowData.Add(item[col].ToString());
                }
                string builtRow = string.Join(",", rowData.ToArray());
                myWriter.WriteLine(builtRow);
            }
            myWriter.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <param name="fileName"></param>
        /// <param name="delim"></param>
        public void GenerateCSV(ref DataTable data, string path, string fileName, string delim)
        {
            System.IO.StreamWriter myWriter = new System.IO.StreamWriter(path + @"\" + fileName);
            StringBuilder row = new StringBuilder();
            foreach (DataColumn item in data.Columns)
            {
                if (string.IsNullOrEmpty(row.ToString()))
                {
                    row.Append(item);
                }
                else
                {
                    row.Append(delim + item.ColumnName);
                }
            }
            myWriter.WriteLine(row);
            foreach (DataRow item in data.Rows)
            {
                List<string> rowData = new List<string>();
                for (int col = 0; col < item.ItemArray.Count(); col++)
                {
                    rowData.Add(item[col].ToString());
                }
                string builtRow = string.Join(delim, rowData.ToArray());
                myWriter.WriteLine(builtRow);
            }
            myWriter.Close();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <param name="prefix"></param>
        /// <param name="maxPerBatch"></param>
        /// <param name="excludeBatchNumber"></param>
        /// <param name="delimiter"></param>
        public string GenerateCSV(ref DataTable data, string path, string prefix, int maxPerBatch, Boolean excludeBatchNumber, string delimiter,string EmptyString)
        {
            string namegen = "";
            int counter = 1;
            string suffix = Environment.UserName + DateTime.Now.ToString("ddMMyyyyHHmm");
            DataColumn col = new DataColumn();
            col.ColumnName = "BatchNumber";
            data.Columns.Add(col);
            int batchRow = 1;
            System.IO.StreamWriter myWriter = null;

            foreach (DataRow item in data.Rows)
            {
                //Create a new file and write the header as required
                if (batchRow == 1)
                {
                    string fileName = prefix + suffix + counter + ".csv";
                    namegen = fileName;
                    myWriter = new System.IO.StreamWriter(path + @"\" + fileName);
                    StringBuilder row = new StringBuilder();
                    foreach (DataColumn headerCol in data.Columns)
                    {
                        if (headerCol.ColumnName == "BatchNumber" && excludeBatchNumber == true)
                        {
                            //Ignore the column
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(row.ToString()))
                            {
                                row.Append(headerCol.ColumnName);
                            }
                            else
                            {
                                row.Append(delimiter + headerCol.ColumnName);
                            }
                        }
                    }
                    myWriter.WriteLine(row);
                }
                List<string> rowData = new List<string>();
                for (int colIndex = 0; colIndex < item.ItemArray.Count(); colIndex++)
                {
                    if (data.Columns[colIndex].ColumnName == "BatchNumber" && excludeBatchNumber == true)
                    {
                        //Ignore the column
                    }
                    else
                    {
                        rowData.Add(item[colIndex].ToString());
                    }
                }
                string builtRow = string.Join(delimiter, rowData.ToArray());
                myWriter.WriteLine(builtRow);
                item.BeginEdit();
                item["BatchNumber"] = prefix + " " + suffix + counter;
                item.EndEdit();
                batchRow++;
                if (batchRow > maxPerBatch)
                {
                    myWriter.Close();
                    batchRow = 1;
                    counter++;
                }
            }
            myWriter.Close();
            return namegen;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <param name="prefix"></param>
        /// <param name="maxPerBatch"></param>
        /// <param name="excludeBatchNumber"></param>
        /// <param name="delimiter"></param>
        /// <param name="ignore"></param>
        /// <param name="addConsec"></param>
        /// <returns></returns>
        public string GenerateCSVRemoveColumn(ref DataTable data, string path, string prefix, int maxPerBatch, Boolean excludeBatchNumber, string delimiter,string ignore,bool addConsec=true)
        {
            string namegen = "";
            int counter = 1;
            string suffix = Environment.UserName + DateTime.Now.ToString("ddMMyyyyHHmm");
            DataColumn col = new DataColumn();
            col.ColumnName = "BatchNumber";
            data.Columns.Add(col);
            int batchRow = 1;
            System.IO.StreamWriter myWriter = null;

            if (addConsec)
            {
                data.Columns.Add("CONSEC");
            }

            foreach (DataRow item in data.Rows)
            {
                //Create a new file and write the header as required
                if (batchRow == 1)
                {
                    string fileName = prefix + suffix + counter + ".csv";
                    namegen = fileName;
                    myWriter = new System.IO.StreamWriter(path + @"\" + fileName);
                    StringBuilder row = new StringBuilder();
                    foreach (DataColumn headerCol in data.Columns)
                    {
                        if ((headerCol.ColumnName == "BatchNumber" && excludeBatchNumber == true)||
                             (headerCol.ColumnName == ignore)|| headerCol.ColumnName=="CONSEC")
                        {
                            //Ignore the column
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(row.ToString()))
                            {
                                row.Append(headerCol.ColumnName);
                            }
                            else
                            {
                                row.Append(delimiter + headerCol.ColumnName);
                            }
                        }
                    }
                    
                    myWriter.WriteLine(row);
                }
                List<string> rowData = new List<string>();
                for (int colIndex = 0; colIndex < item.ItemArray.Count(); colIndex++)
                {
                    if ((data.Columns[colIndex].ColumnName == "BatchNumber" && excludeBatchNumber == true) || 
                        data.Columns[colIndex].ColumnName ==ignore || data.Columns[colIndex].ColumnName == "CONSEC")
                    {
                        //Ignore the column
                    }
                    else
                    {
                        rowData.Add(item[colIndex].ToString());
                    }
                }
                string builtRow = string.Join(delimiter, rowData.ToArray());
                myWriter.WriteLine(builtRow);
                item.BeginEdit();
                item["BatchNumber"] = prefix + " " + suffix + counter;
                if (addConsec)
                {
                    item["CONSEC"] = counter;
                }
                item.EndEdit();
                batchRow++;
                if (batchRow > maxPerBatch)
                {
                    myWriter.Close();
                    batchRow = 1;
                    counter++;
                }
            }
            myWriter.Close();
            return namegen;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        /// <param name="prefix"></param>
        /// <param name="maxPerBatch"></param>
        /// <param name="excludeBatchNumber"></param>
        public void GenerateCSV(ref DataTable data, string path, string prefix, int maxPerBatch, Boolean excludeBatchNumber)
        {
            GenerateCSV(ref data, path, prefix, maxPerBatch, excludeBatchNumber, ",",string.Empty);
        }

        #endregion

        #region Transform data show row as table
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowToConvert"></param>
        /// <param name="headerForFieldColumn"></param>
        /// <param name="headerForValueColumn"></param>
        /// <returns></returns>
        public DataTable ConvertRowToTable(DataRow rowToConvert, string headerForFieldColumn, string headerForValueColumn)
        {
            DataTable temp = new DataTable();
            DataColumn col = new DataColumn(headerForFieldColumn);
            temp.Columns.Add(col);
            col = new DataColumn(headerForValueColumn);
            temp.Columns.Add(col);
            foreach(DataColumn item in rowToConvert.Table.Columns)
            {
                DataRow newRow = temp.NewRow();
                newRow[headerForFieldColumn] = item.ColumnName;
                newRow[headerForValueColumn] = rowToConvert[item.ColumnName].ToString();
                temp.Rows.Add(newRow);
            }
            return temp;
        }

        #endregion
        /// <summary>
        /// 
        /// </summary>
        public static OleDbConnection m_connection;
        /// <summary>
        /// 
        /// </summary>
        public static SqlConnection m_SQLconnection;
        OleDbDataAdapter m_localAdapter;
        SqlDataAdapter m_localSQLAdapter;
        string dataType = "";

        /// <summary>
        /// An adapter containing to the connection to the data source
        /// </summary>
        public OleDbDataAdapter localAdapter
        {
            get
            {
                return m_localAdapter;
            }
        }

        /// <summary>
        /// An adapter containing the connection to the data source
        /// </summary>
        public SqlDataAdapter SQLAdapter
        {
            get
            {
                return m_localSQLAdapter;
            }
        }

        /// <summary>
        /// Creates a default instance of the class - there will be no disconnected access to the underlying
        /// data source
        /// </summary>
        public DataHandler()
        {
        }

        /// <summary>
        /// Creates an instance of the class. As part of creating the instance the class connects to the data
        /// source
        /// </summary>
        /// <param name="connectionString">The provider string for connecting to the data source</param>
        /// <param name="queryString">The query to be applied when creating the connection</param>
        /// <remarks>This version of the method requires the full provider string (including any relevant password). It is therefore not
        /// recommended for use with password protected data sources. Use the overloaded method which allows you to pass an encrypted
        /// password when initializing the connection</remarks>
        /// <example>
        /// This example shows you how to connect to a non-password protected data source
        /// <code>
        /// public void ConnectToResultsData()
        /// {
        ///    string provider = Provider=Microsoft.Jet.OLEDB.4.0;Data Source="C:\SomeFolder\MyProject\Data.mdb";Jet OLEDB:Database Password=";
        ///    ResultsHandler = new New_Wrapper.DataHandler(provider, "SELECT Results.CaseId, Results.Result FROM Results");
        /// }
        /// </code>
        ///</example>
        public DataHandler(string connectionString, string queryString)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (connectionString.IndexOf("Server=") == -1)
            {
                // Access connection string
                dataType = "Access";
                CreateAdapter(connectionString, queryString, ref m_localAdapter);
            }
            else
            {
                // SQL connection string
                dataType = "SQL";
                CreateAdapter(connectionString, queryString, ref m_localSQLAdapter);
            }
        }

        /// <summary>
        /// Connects to a database with an encrypted password
        /// </summary>
        /// <param name="connectionString">The base connection string without the password</param>
        /// <param name="password">The encrypted password</param>
        /// <param name="queryString">The query to be applied when creating the connection</param>
        /// <remarks>This allows you to pass an encrypted password to be used when connecting to the data source. The password is
        /// decrypted by the DMB development teams encryption class</remarks>
        /// <example>This example shows how to connect to a password protected database
        /// <code>
        /// public void ConnectToResultsData()
        /// {
        ///    string provider = Provider=Microsoft.Jet.OLEDB.4.0;Data Source="C:\SomeFolder\MyProject\Data.mdb";Jet OLEDB:Database Password=";
        ///    string password = "thisisencrypted";
        ///    ResultsHandler = new New_Wrapper.DataHandler(provider, password, "SELECT Results.CaseId, Results.Result FROM Results");
        /// }
        /// </code>
        /// </example>
        //public DataHandler(string connectionString, string password, string queryString)
        //{
        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

        //    // ### ...and build up the full connection string ###
        //    connectionString = connectionString + decryptedPassword + ";";

        //    // ### As a password is involved, the data source must be Access ###
        //    dataType = "Access";

        //    CreateAdapter(connectionString, queryString, ref m_localAdapter);
        //}

        /// <summary>
        /// Creates a data adapter to handle communication between the class and the data source
        /// </summary>
        /// <param name="passedConnString">The provider string for connecting to the data source</param>
        /// <param name="passedQuery">The query to be applied when creating the connection</param>
        /// <param name="passedAdapter">The adapter that the connection will stored into</param>
        private void CreateAdapter(string passedConnString, string passedQuery, ref OleDbDataAdapter passedAdapter)
        {
            m_connection = new OleDbConnection(passedConnString);
            m_connection.Open();
            passedAdapter = new OleDbDataAdapter();
            passedAdapter.SelectCommand = new OleDbCommand(passedQuery, m_connection);
            OleDbCommandBuilder builder = new OleDbCommandBuilder(passedAdapter);
        }

        // ### SQL version of above ###
        private void CreateAdapter(string passedConnString, string passedQuery, ref SqlDataAdapter passedAdapter)
        {
            m_SQLconnection = new SqlConnection(passedConnString);
            m_SQLconnection.Open();
            passedAdapter = new SqlDataAdapter();
            passedAdapter.SelectCommand = new SqlCommand(passedQuery, m_SQLconnection);
            SqlCommandBuilder builder = new SqlCommandBuilder(passedAdapter);
        }

        /// <summary>
        /// Creates a dataset based on the data source connection
        /// </summary>
        /// <returns>A dataset containing the filtered data from the data source</returns>
        /// <example>
        /// The method is used after successfully conneccting to the datasource to populate a dataset with the resulting data
        /// <code>
        /// public void ConnectToResultsData()
        /// {
        ///    string provider = Provider=Microsoft.Jet.OLEDB.4.0;Data Source="C:\SomeFolder\MyProject\Data.mdb";Jet OLEDB:Database Password=";
        ///    string password = "thisisencrypted";
        ///    ResultsHandler = new New_Wrapper.DataHandler(provider, password, "SELECT Results.CaseId, Results.Result FROM Results");
        ///    ExistingReturns = ResultsHandler.CreateDataset();
        /// }
        /// </code>
        /// </example>
        public DataSet CreateDataset()
        {
            DataSet dataSet = new DataSet();
            if (dataType == "SQL")
            {
                m_localSQLAdapter.Fill(dataSet);
            }
            else
            {
                // ### Must be Access ###
                m_localAdapter.Fill(dataSet);
            }
            return dataSet;
        }

        /// <summary>
        /// Updates the data source with any changes made to the dataset
        /// </summary>
        /// <param name="passedDataset">The dataset that contains the changes</param>
        /// <example>To store any changes made to the dataset you simply need to pass the dataset to the method. Any entries that are
        /// flagged as having been changed will be committed to the datasource (this includes, additions, edits and deletions)
        /// <code>
        ///  //Your code to update the dataset
        ///  myDataHandler.UpdateDataset(myDataset);
        /// </code>
        /// Note - the dataset being updated must be the one created from this particular instance of the datahandler - trying to pass a
        /// dataset not generated by the specific instance of the class will result in a data exception
        /// </example>
        public void UpdateDataset(DataSet passedDataset)
        {
            if (dataType == "SQL")
            {
                SqlCommandBuilder builder = new SqlCommandBuilder(m_localSQLAdapter);
                m_localSQLAdapter.ContinueUpdateOnError = true;
                m_localSQLAdapter.UpdateCommand = builder.GetUpdateCommand();
                m_localSQLAdapter.Update(passedDataset, passedDataset.Tables[0].TableName);
            }
            else
            {
                // ### Must be Access ###
                OleDbCommandBuilder builder = new OleDbCommandBuilder(m_localAdapter);
                builder.QuotePrefix = "[";
                builder.QuoteSuffix = "]";
                m_localSQLAdapter.ContinueUpdateOnError = true;
                m_localAdapter.UpdateCommand = builder.GetUpdateCommand();
                m_localAdapter.Update(passedDataset, passedDataset.Tables[0].TableName);
            }
        }

        /// <summary>
        /// Creates a DataSet based on a SQL string
        /// This method creates a separate data adapter to process the data request - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="getDataQuery">The query to be used when creating the DataSet based on the data source</param>
        public DataSet GetDataSet(string provider, string getDataQuery)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (provider.IndexOf("Server=") == -1)
            {
                // Must be Access
                try
                {
                    DataSet getData = new DataSet();
                    OleDbConnection newConnection = new OleDbConnection(provider);
                    newConnection.Open();
                    OleDbDataAdapter newAdapter = new OleDbDataAdapter(getDataQuery, newConnection);
                    newAdapter.Fill(getData, "Temp");
                    return getData;
                }
                catch (Exception ex)
                {
                    throw new Exception("The DataSet could not be created for the following reason - " + ex.Message);
                }
            }
            else
            {
                // Must be SQL
                try
                {
                    DataSet getData = new DataSet();
                    SqlConnection newConnection = new SqlConnection(provider);
                    newConnection.Open();
                    SqlDataAdapter newAdapter = new SqlDataAdapter(getDataQuery, newConnection);
                    newAdapter.Fill(getData, "Temp");
                    return getData;
                }
                catch (Exception ex)
                {
                    throw new Exception("The DataSet could not be created for the following reason - " + ex.Message);
                }
            }
            
        }

        /// <summary>
        /// Creates a DataSet based on a SQL string
        /// This method creates a separate data adapter to process the data request - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="getDataQuery">The query to be used when creating the DataSet based on the data source</param>
        //public DataSet GetDataSet(string provider, string password, string getDataQuery)
        //{
        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

        //    // ### ...and build up the full connection string ###
        //    provider = provider + decryptedPassword + ";";

        //    try
        //    {
        //        DataSet getData = new DataSet();
        //        OleDbConnection newConnection = new OleDbConnection(provider);
        //        newConnection.Open();
        //        OleDbDataAdapter newAdapter = new OleDbDataAdapter(getDataQuery, newConnection);
        //        newAdapter.Fill(getData, "Temp");
        //        return getData;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("The DataSet could not be created for the following reason - " + ex.Message);
        //    }
        //}

        /// <summary>
        /// Inserts data directly into the data source
        /// This method creates a separate data adapter to process the insert command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="insertQuery">The insert query to be used when updating the data source</param>
        public void InsertData(string provider, string insertQuery)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (provider.IndexOf("Server=") == -1)
            {
                // Must be Access
                try
                {
                    OleDbDataAdapter newAdapter = new OleDbDataAdapter();
                    OleDbConnection newConnection = new OleDbConnection(provider);
                    newConnection.Open();
                    newAdapter.InsertCommand = new OleDbCommand(insertQuery, newConnection);
                    newAdapter.InsertCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception("The data could not be inserted for the following reason - " + ex.Message);
                }
            }
            else
            {
                // Must be SQL
                try
                {
                    SqlDataAdapter newAdapter = new SqlDataAdapter();
                    SqlConnection newConnection = new SqlConnection(provider);
                    newConnection.Open();
                    newAdapter.InsertCommand = new SqlCommand(insertQuery, newConnection);
                    newAdapter.InsertCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception("The data could not be inserted for the following reason - " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Inserts data directly into the data source
        /// This method creates a separate data adapter to process the insert command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="insertQuery">The insert query to be used when updating the data source</param>
        //public void InsertData(string provider, string password, string insertQuery)
        //{
        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

        //    // ### ...and build up the full connection string ###
        //    provider = provider + decryptedPassword + ";";

        //    try
        //    {
        //        OleDbDataAdapter newAdapter = new OleDbDataAdapter();
        //        OleDbConnection newConnection = new OleDbConnection(provider);
        //        newConnection.Open();
        //        newAdapter.InsertCommand = new OleDbCommand(insertQuery, newConnection);
        //        newAdapter.InsertCommand.ExecuteNonQuery();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("The data could not be inserted for the following reason - " + ex.Message);
        //    }
        //}

        /// <summary>
        /// Deletes data directly from the data source
        /// this method creates a separate data adapter to process the delete command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="deleteQuery">The delete query to be used when updating the data source</param>
        public void DeleteData(string provider, string deleteQuery)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (provider.IndexOf("Server=") == -1)
            {
                // Must be Access
                try
                {
                    OleDbDataAdapter newAdapter = new OleDbDataAdapter();
                    OleDbConnection newConnection = new OleDbConnection(provider);
                    newConnection.Open();
                    newAdapter.DeleteCommand = new OleDbCommand(deleteQuery, newConnection);
                    newAdapter.DeleteCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception("The data could not be deleted for the following reason - " + ex.Message);
                }
            }
            else
            {
                // Must be SQL
                try
                {
                    SqlDataAdapter newAdapter = new SqlDataAdapter();
                    SqlConnection newConnection = new SqlConnection(provider);
                    newConnection.Open();
                    newAdapter.DeleteCommand = new SqlCommand(deleteQuery, newConnection);
                    newAdapter.DeleteCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception("The data could not be deleted for the following reason - " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Deletes data directly from the data source
        /// this method creates a separate data adapter to process the delete command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="deleteQuery">The delete query to be used when updating the data source</param>
        //public void DeleteData(string provider, string password, string deleteQuery)
        //{
        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

        //    // ### ...and build up the full connection string ###
        //    provider = provider + decryptedPassword + ";";

        //    try
        //    {
        //        OleDbDataAdapter newAdapter = new OleDbDataAdapter();
        //        OleDbConnection newConnection = new OleDbConnection(provider);
        //        newConnection.Open();
        //        newAdapter.DeleteCommand = new OleDbCommand(deleteQuery, newConnection);
        //        newAdapter.DeleteCommand.ExecuteNonQuery();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("The data could not be deleted for the following reason - " + ex.Message);
        //    }
        //}

        /// <summary>
        /// Updates data directly from the data source
        /// this method creates a separate data adapter to process the update command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="updateQuery">The update query to be used when updating the data source</param>
        public void UpdateData(string provider, string updateQuery)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (provider.IndexOf("Server=") == -1)
            {
                // Must be Access
                try
                {
                    dataType = "ADO";
                    OleDbDataAdapter newAdapter = new OleDbDataAdapter();
                    OleDbConnection newConnection = new OleDbConnection(provider);
                    newConnection.Open();
                    newAdapter.UpdateCommand = new OleDbCommand(updateQuery, newConnection);
                    newAdapter.UpdateCommand.ExecuteNonQuery();
                    newConnection.Close();
                }
                catch (Exception ex)
                {
                    throw new Exception("The data could not be updated for the following reason - " + ex.Message);
                }
            }
            else
            {
                // Must be SQL
                try
                {
                    dataType = "SQL";
                    SqlDataAdapter newAdapter = new SqlDataAdapter();
                    SqlConnection newConnection = new SqlConnection(provider);
                    newConnection.Open();
                    newAdapter.UpdateCommand = new SqlCommand(updateQuery, newConnection);
                    newAdapter.UpdateCommand.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception("The data could not be updated for the following reason - " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Updates data directly from the data source
        /// this method creates a separate data adapter to process the update command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="updateQuery">The query for updating the data source</param>
        //public void UpdateData(string provider, string password, string updateQuery)
        //{
        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

        //    // ### ...and build up the full connection string ###
        //    provider = provider + decryptedPassword + ";";
            
        //    try
        //    {
        //        OleDbDataAdapter newAdapter = new OleDbDataAdapter();
        //        OleDbConnection newConnection = new OleDbConnection(provider);
        //        newConnection.Open();
        //        newAdapter.UpdateCommand = new OleDbCommand(updateQuery, newConnection);
        //        newAdapter.UpdateCommand.ExecuteNonQuery();
        //        newConnection.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("The data could not be updated for the following reason - " + ex.Message);
        //    }
        //}

        /// <summary>
        /// 
        /// </summary>
        public void CloseConnection()
        {
            if (dataType == "SQL")
            {
                m_SQLconnection.Close();
            }
            else
            {
                m_connection.Close();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dispose"></param>

        public void CloseConnection(bool dispose)
        {
            if (dispose == false)
            {
                CloseConnection();
                return;
            }
            if (dataType == "SQL")
            {
                m_SQLconnection.Dispose();
            }
            else
            {
                m_connection.Dispose();
            }
        }

        /// <summary>
        /// Close the connection based on what type of connection has been made
        /// </summary>
        /// <param name="dataType">The data engine that was used to connectio to the data source (i.e. SQL, ADO)</param>
        public void CloseConnection(string dataType)
        {
            if (dataType == "SQL")
            {
                m_SQLconnection.Close();
            }
            else
            {
                m_connection.Close();
            }
        }
        /// <summary>
        /// Populates a dataset with multiple tables
        /// </summary>
        /// <param name="queryList">The tables to be inckuded in the dataset. Format = [query string; table name]</param>
        /// <returns>A dataset containing the filtered data from the data source</returns>
        /// <remarks>As you are changing the Select command for the data adapter this method should only be used when viewing
        /// statuc data such as look up tables</remarks>
        public DataSet CreateDataset(string[] queryList)
        {
            DataSet dataSet = new DataSet();
            if (dataType == "SQL")
            {
                foreach (string item in queryList)
                {
                    string[] fields = item.Split(';');
                    m_localSQLAdapter.SelectCommand.CommandText = fields[0];
                    m_localSQLAdapter.Fill(dataSet, fields[1]);
                    dataSet.Tables[dataSet.Tables.Count - 1].TableName = fields[1];
                }
            }
            else
            {
                // ### Must be Access ###
                foreach (string item in queryList)
                {
                    string[] fields = item.Split(';');
                    m_localAdapter.SelectCommand.CommandText = fields[0];
                    m_localAdapter.Fill(dataSet, fields[1]);
                    dataSet.Tables[dataSet.Tables.Count - 1].TableName = fields[1];
                }
            }
            return dataSet;
        }

        #region "Schema data"

        /// <summary>
        /// Gets a list of the 'databases' mounted on a specified SQL instance
        /// </summary>
        /// <param name="serverName">The name of the SQL server</param>
        /// <returns>A data table containing the schema data for the 'databases'</returns>
        public DataTable RetrieveListOfSQLDatabases(string serverName)
        {
            string connectionString = "Server = " + serverName + ";Integrated Security = sspi;";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlDataAdapter adapter = new SqlDataAdapter("SELECT Name FROM master.sys.databases", connection))
                {
                    DataSet temp = new DataSet();
                    adapter.Fill(temp);
                    return temp.Tables[0];
                }
            }
        }

        /// <summary>
        /// Gets a list of tables in a specified SQL database
        /// </summary>
        /// <param name="serverName">The name of the server that the database is mounted on</param>
        /// <param name="databaseName">The name of the database</param>
        /// <returns>A data table containing the tables schema for the database</returns>
        public DataTable RetrieveListOfTablesInSQLDatabase(string serverName, string databaseName)
        {
            string connectionString = @"Data Source = " + serverName + ";Database=" + databaseName + ";Integrated Security=true;";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                DataTable myTable = connection.GetSchema("Tables");
                return myTable;
            }
        }

        /// <summary>
        /// Retrieve a list of the columns in a specified SQL table
        /// </summary>
        /// <param name="serverName">The name of the server that the database is mounted on</param>
        /// <param name="databaseName">The name of the database</param>
        /// <param name="tableName">The name of the table in the database</param>
        /// <returns>A data table containing the columns schema for the specified table</returns>
        public DataTable RetrieveColumnSchemaForSQLTable(string serverName, string databaseName, string tableName)
        {
            string connectionString = @"Data Source = " + serverName + ";Database=" + databaseName + ";Integrated Security=true;";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string[] restrictions = new string[4];
                restrictions[2] = tableName;
                DataTable myTable = connection.GetSchema("Columns", restrictions);
                return myTable;
            }
        }

        /// <summary>
        /// Retrieve a list of the tables in a specified Access database
        /// </summary>
        /// <param name="filePath">The full path to the mdb file (including path and file name)</param>
        /// <param name="password">The password for the database</param>
        /// <param name="encrypted">Whether the password is encrypted</param>
        /// <returns>A data table containing the tables schema for the specified database</returns>
        //public DataTable RetrieveListOfTablesInAccessDatabase(string filePath, string password, Boolean encrypted)
        //{
        //    if (encrypted == true)
        //    {
        //        password = EncryptString.StringCipher.Decrypt(password, "SkimmedMilk");
        //    }
        //    string provider = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + filePath + "';Jet OLEDB:Database Password=" + password;

        //    System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection(provider);
        //    connection.Open();
        //    string[] restrictions = new string[4];
        //    restrictions[3] = "TABLE";
        //    return connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, restrictions);
        //}

        /// <summary>
        /// Retrieve a list of the columns in a specified Access table
        /// </summary>
        /// <param name="filePath">The full path to the mdb file (including path and file name)</param>
        /// <param name="password">The password for the database</param>
        /// <param name="encrypted">Whether the password is encrypted</param>
        /// <param name="tableName">The name of the Access data table</param>
        /// <returns>A data table containing the columns schema for the specified table</returns>
        //public DataTable RetrieveColumnSchemaForAccessTable(string filePath, string password, Boolean encrypted, string tableName)
        //{
        //    if (encrypted == true)
        //    {
        //        password = EncryptString.StringCipher.Decrypt(password, "SkimmedMilk");
        //    }
        //    string provider = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + filePath + "';Jet OLEDB:Database Password=" + password;
        //    System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection(provider);
        //    connection.Open();
        //    string[] restrictions = new string[4];
        //    restrictions[2] = tableName;
        //    DataTable myTable = connection.GetSchema("Columns", restrictions);
        //    return myTable;
        //}

        #endregion

    }

    /// <summary>
    /// Handles the connecting to and updating of data using OLE
    /// </summary>
    /// <remarks>The MultiTableHandler allows you to update multiple tables within a single dataset</remarks>
    public class MultiTableHandler : New_Wrapper.DataHandler
    {

        #region "Fields"

        //OleDbConnection m_connection;
        OleDbDataAdapter m_localAdapter;
        //SqlConnection m_SQLconnection;
        SqlDataAdapter m_localSQLAdapter;
        string dataType = "";

        #endregion

        #region "Properties"

        string[,] m_selectQueries;
        /// <summary>
        /// An multi-dimensional array of the queries that will be used to retrieve data from the data source
        /// </summary>
        /// <remarks>The table name should be the first 'sub element' followed by the appropriate SQL query as
        /// the second 'sub element'
        /// For instance, selectQueries[0,0] = "tblMain" selectQueries[0,1] = "SELECT...."</remarks>
        public string[,] selectQueries
        {
            get { return m_selectQueries; }
            set { m_selectQueries = value; }
        }

        string m_providerString = "";
        /// <summary>
        /// The string for connecting to the data source
        /// </summary>
        public string providerString
        {
            get { return m_providerString; }
            set { m_providerString = value; }
        }

        DataSet m_sourceData = new DataSet();
        /// <summary>
        /// A dataset containing the data tables
        /// </summary>
        public DataSet sourceData
        {
            get { return m_sourceData; }
            set { m_sourceData = value; }
        }

        #endregion

        #region "Class initialisation"

        /// <summary>
        /// Default intialiser for the class. 
        /// </summary>
        /// <remarks>The properties for the class must be set prior to using the class to update or retrieve
        /// data</remarks>
        public MultiTableHandler()
        {
        }

        /// <summary>
        /// Initialises the class and set the relevant properties
        /// </summary>
        /// <param name="passedQueries">The array containing the list of tables and associated queries</param>
        /// <param name="passedProvider">The provider string for connecting to the database</param>
        /// <remarks>The table name should be the first 'sub element' followed by the appropriate SQL query as
        /// the second 'sub element'
        /// For instance, selectQueries[0,0] = "tblMain" selectQueries[0,1] = "SELECT...."</remarks>
        public MultiTableHandler(string[,] passedQueries, string passedProvider)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (passedProvider.Substring(0, 8) == "Provider")
            {
                dataType = "Access";
            }
            else
            {
                dataType = "SQL";
            }
            selectQueries = passedQueries;
            providerString = passedProvider;
            connectToData();
        }

        /// <summary>
        /// Initialises the class and set the relevant properties
        /// </summary>
        /// <param name="passedQueries">The array containing the list of tables and associated queries</param>
        /// <param name="passedProvider">The provider sting for connecting to the database</param>
        /// <param name="passedPassword">The eencrypted password of the database</param>
        /// <remarks>The table name should be the first 'sub element' followed by the appropriate SQL query as
        /// the second 'sub element'
        /// For instance, selectQueries[0,0] = "tblMain" selectQueries[0,1] = "SELECT...."</remarks>
        //public MultiTableHandler(string[,] passedQueries, string passedProvider, string passedPassword)
        //{
        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(passedPassword, "SkimmedMilk");

        //    // ### ...and build up the full connection string ###
        //    passedProvider = passedProvider + decryptedPassword + ";";

        //    // ### Password has been passed so the data type must be Access
        //    dataType = "Access";

        //    selectQueries = passedQueries;
        //    providerString = passedProvider;
        //    connectToData();
        //}

        #endregion

        #region "Public methods"

        /// <summary>
        /// Connects to the data source and populates a data set with the requested tables
        /// </summary>
        /// <param name="passedQueries">The table names and SQL queries that the data tables will be based on</param>
        /// <param name="passedProvider">The provider string for connecting to the data source</param>
        public void connectToData(string[,] passedQueries, string passedProvider)
        {
            // ### No password being passed, check the connection string for the word "Provider" ###
            if (passedProvider.Substring(0, 8) == "Provider")
            {
                dataType = "Access";
            }
            else
            {
                dataType = "SQL";
            }

            selectQueries = passedQueries;
            providerString = passedProvider;
            connectToData();
        }

        /// <summary>
        /// Connects to the data source and populates a data set with the requested tables
        /// </summary>
        /// <param name="passedQueries">The table names and SQL queries that the data tables will be based on</param>
        /// <param name="passedProvider">The provider string for connecting to the data source</param>
        /// <param name="passedPassword">The database's encrypted password</param>
        //public void connectToData(string[,] passedQueries, string passedProvider, string passedPassword)
        //{
        //    // ### Password has been passed so the data type must be Access
        //    dataType = "Access";

        //    // ### Decrypt the encrypted password... ###
        //    string decryptedPassword = StringCipher.Decrypt(passedPassword, "SkimmedMilk");
        //    selectQueries = passedQueries;
        //    providerString = passedProvider + decryptedPassword + ";";
        //    connectToData();
        //}

        /// <summary>
        /// Connects to the data source and populates a data set with the requested tables
        /// </summary>
        public void connectToData()
        {
            if (dataType == "SQL")
            {
                if (selectQueries == null || selectQueries.GetLength(0) == 0)
                {
                    throw new Exception("There are no table names or queries on which to base the connection.");
                }
                if (providerString == "")
                {
                    throw new Exception("The provider string for connecting to the data has not been initialised.");
                }
                m_SQLconnection = new SqlConnection(providerString);
                try
                {
                    m_SQLconnection.Open();
                    m_localSQLAdapter = new SqlDataAdapter();
                }
                catch
                {
                    throw new Exception("Unable to connect to the data source.");
                }
                try
                {
                    for (int i = 0; i <= selectQueries.GetLength(0) - 1; i++)
                    {
                        m_localSQLAdapter.SelectCommand = new SqlCommand(selectQueries[i, 1], m_SQLconnection);
                        m_localSQLAdapter.Fill(m_sourceData, selectQueries[i, 0]);
                        m_sourceData.Tables[i].TableName = selectQueries[i, 0];
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(@"Unable to resolve tables\SQL queries - " + ex.Message);
                }
            }
            else
            {
                // Must be Access
                if (selectQueries == null || selectQueries.GetLength(0) == 0)
                {
                    throw new Exception("There are no table names or queries on which to base the connection.");
                }
                if (providerString == "")
                {
                    throw new Exception("The provider string for connecting to the data has not been initialised.");
                }
                m_connection = new OleDbConnection(providerString);
                try
                {
                    m_connection.Open();
                    m_localAdapter = new OleDbDataAdapter();
                }
                catch
                {
                    throw new Exception("Unable to connect to the data source.");
                }
                try
                {
                    for (int i = 0; i <= selectQueries.GetLength(0) - 1; i++)
                    {
                        m_localAdapter.SelectCommand = new OleDbCommand(selectQueries[i, 1], m_connection);
                        m_localAdapter.Fill(m_sourceData, selectQueries[i, 0]);
                        m_sourceData.Tables[i].TableName = selectQueries[i, 0];
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(@"Unable to resolve tables\SQL queries - " + ex.Message);
                }
            }
        }

        /// <summary>
        /// Updates the specified table(s) in the data set
        /// </summary>
        /// <param name="tableName">The name of the table that is to be updated</param>
        public void UpdateData(string tableName)
        {
            if (dataType == "SQL")
            {
                SqlDataAdapter tempAdapter = new SqlDataAdapter();
                SqlCommandBuilder tempBuilder = new SqlCommandBuilder(tempAdapter);
                tempAdapter.AcceptChangesDuringFill = false;
                tempAdapter.SelectCommand = new SqlCommand("SELECT * FROM " + tableName, m_SQLconnection);
                tempAdapter.InsertCommand = tempBuilder.GetInsertCommand();
                tempAdapter.UpdateCommand = tempBuilder.GetUpdateCommand();
                tempAdapter.TableMappings.Add("Table", tableName);
                DataSet ds = new DataSet();
                DataTable tempTable = sourceData.Tables[tableName].Copy();
                tempTable.TableName = tableName;
                ds.Merge(tempTable);
                tempAdapter.Update(ds, tableName);
                sourceData.Tables[tableName].AcceptChanges();
            }
            else
            {
                // Must be Access
                OleDbDataAdapter tempAdapter = new OleDbDataAdapter();
                OleDbCommandBuilder tempBuilder = new OleDbCommandBuilder(tempAdapter);
                tempAdapter.AcceptChangesDuringFill = false;
                tempAdapter.SelectCommand = new OleDbCommand("SELECT * FROM " + tableName, m_connection);
                tempAdapter.InsertCommand = tempBuilder.GetInsertCommand();
                tempAdapter.UpdateCommand = tempBuilder.GetUpdateCommand();
                tempAdapter.TableMappings.Add("Table", tableName);
                DataSet ds = new DataSet();
                //tempAdapter.Fill(ds);
                DataTable tempTable = sourceData.Tables[tableName].Copy();
                tempTable.TableName = tableName;
                ds.Merge(tempTable);
                tempAdapter.Update(ds, tableName);
                sourceData.Tables[tableName].AcceptChanges();
            }
        }

        /// <summary>
        /// Adds a table from the multihandler instance to an external dataset
        /// </summary>
        /// <param name="passedDataset">The dataset that the table is to be copied to</param>
        /// <param name="passedTableName">The name of the table to be copied to the external dataset</param>
        /// <remarks>The table being copied must already exist in the sourceData property of the class</remarks>
        public void AddTableToDataset(ref DataSet passedDataset, string passedTableName)
        {
            if (dataType == "SQL")
            {
                m_localSQLAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                DataTable myTable = sourceData.Tables[passedTableName].Clone();
                passedDataset.Tables.Add(myTable);
                passedDataset.Tables[passedTableName].Merge(sourceData.Tables[passedTableName]);
            }
            else
            {
                // Must be Access
                m_localAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                //m_localAdapter.Fill(passedDataset, passedTableName);
                DataTable myTable = sourceData.Tables[passedTableName].Clone();
                //sourceData.Tables[passedTableName].Merge();
                passedDataset.Tables.Add(myTable);
                passedDataset.Tables[passedTableName].Merge(sourceData.Tables[passedTableName]);
            }
        }

        /// <summary>
        /// Adds a table from the multihandler instance to an external dataset
        /// </summary>
        /// <param name="passedDataset">The dataset that the table is to be copied to</param>
        /// <param name="passedTableName">The name of the table to be copied to the external dataset</param>
        /// <param name="filter">The filter that is to be applied to the data</param>
        /// <param name="passedSort">Any sort that is to be applied to the data</param>
        public void AddTableToDataset(ref DataSet passedDataset, string passedTableName, string filter, string passedSort)
        {
            if (dataType == "SQL")
            {
                m_localSQLAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                DataTable myTable = sourceData.Tables[passedTableName].Clone();

                DataView localView = new DataView(sourceData.Tables[passedTableName], filter, "", DataViewRowState.CurrentRows);
                foreach (DataRowView myRow in localView)
                {
                    DataRow tempRow = myRow.Row;
                    myTable.ImportRow(tempRow);
                }
                passedDataset.Tables.Add(myTable);
            }
            else
            {
                // Must be Access
                m_localAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
                DataTable myTable = sourceData.Tables[passedTableName].Clone();

                DataView localView = new DataView(sourceData.Tables[passedTableName], filter, "", DataViewRowState.CurrentRows);
                foreach (DataRowView myRow in localView)
                {
                    DataRow tempRow = myRow.Row;
                    myTable.ImportRow(tempRow);
                }
                passedDataset.Tables.Add(myTable);
            }
        }

        #endregion
    }
}
