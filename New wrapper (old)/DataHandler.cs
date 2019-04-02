using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using EncryptString;

namespace New_Wrapper
{

    /// <summary>
    /// Handles the connecting to and updating of data using OLE
    /// </summary>
    /// //<remarks>The default instance of the data handler only connects and updates a single table. For datasets containing
    /// multiple updateable tables use the MultiHandler</remarks>
    public class DataHandler
    {

        OleDbConnection m_connection;
        OleDbDataAdapter m_localAdapter;
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
            CreateAdapter(connectionString, queryString, ref m_localAdapter);
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
        public DataHandler(string connectionString, string password, string queryString)
        {
            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

            // ### ...and build up the full connection string ###
            connectionString = connectionString + decryptedPassword + ";";

            CreateAdapter(connectionString, queryString, ref m_localAdapter);
        }


        /// <summary>
        /// Creates a data adapter to handle communication between the class and the data source
        /// </summary>
        /// <param name="passedConnString">The provider string for connecting to the data source</param>
        /// <param name="passedQuery">The query to be applied when creating the connection</param>
        /// <param name="passedAdapter">The adapter that the connection will stored into</param>
        private void CreateAdapter(string passedConnString, string passedQuery,ref OleDbDataAdapter passedAdapter)
        {
            m_connection = new OleDbConnection(passedConnString);
            m_connection.Open();
            passedAdapter = new OleDbDataAdapter();
            passedAdapter.SelectCommand = new OleDbCommand(passedQuery, m_connection);
            OleDbCommandBuilder builder = new OleDbCommandBuilder(passedAdapter);
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
            m_localAdapter.Fill(dataSet);
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
            OleDbCommandBuilder builder = new OleDbCommandBuilder(m_localAdapter);
            m_localAdapter.UpdateCommand = builder.GetUpdateCommand();
            m_localAdapter.Update(passedDataset,passedDataset.Tables[0].TableName);

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
            try
            {
                DataSet getData = new DataSet();
                OleDbConnection newConnection = new OleDbConnection(provider);
                newConnection.Open();
                OleDbDataAdapter newAdapter = new OleDbDataAdapter(getDataQuery, newConnection);
                newAdapter.Fill(getData, "Temp");
                return getData;
            }
            catch(Exception ex)
            {
                throw new Exception("The DataSet could not be created for the following reason - " + ex.Message);
            }
        }

        /// <summary>
        /// Creates a DataSet based on a SQL string
        /// This method creates a separate data adapter to process the data request - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="insertQuery">The query to be used when creating the DataSet based on the data source</param>
        public DataSet GetDataSet(string provider, string password, string getDataQuery)
        {
            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

            // ### ...and build up the full connection string ###
            provider = provider + decryptedPassword + ";";

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

        /// <summary>
        /// Inserts data directly into the data source
        /// This method creates a separate data adapter to process the insert command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="insertQuery">The insert query to be used when updating the data source</param>
        public void InsertData(string provider, string insertQuery)
        {
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

        /// <summary>
        /// Inserts data directly into the data source
        /// This method creates a separate data adapter to process the insert command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="insertQuery">The insert query to be used when updating the data source</param>
        public void InsertData(string provider, string password, string insertQuery)
        {
            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

            // ### ...and build up the full connection string ###
            provider = provider + decryptedPassword + ";";

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

        /// <summary>
        /// Deletes data directly from the data source
        /// this method creates a separate data adapter to process the delete command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="deleteQuery">The delete query to be used when updating the data source</param>
        public void DeleteData(string provider, string deleteQuery)
        {
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

        /// <summary>
        /// Deletes data directly from the data source
        /// this method creates a separate data adapter to process the delete command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="deleteQuery">The delete query to be used when updating the data source</param>
        public void DeleteData(string provider, string password, string deleteQuery)
        {
            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

            // ### ...and build up the full connection string ###
            provider = provider + decryptedPassword + ";";

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

        /// <summary>
        /// Updates data directly from the data source
        /// this method creates a separate data adapter to process the update command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="deleteQuery">The update query to be used when updating the data source</param>
        public void UpdateData(string provider, string updateQuery)
        {
            try
            {
                OleDbDataAdapter newAdapter = new OleDbDataAdapter();
                OleDbConnection newConnection = new OleDbConnection(provider);
                newConnection.Open();
                newAdapter.UpdateCommand = new OleDbCommand(updateQuery, newConnection);
                newAdapter.UpdateCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception("The data could not be updated for the following reason - " + ex.Message);
            }
        }

        /// <summary>
        /// Updates data directly from the data source
        /// this method creates a separate data adapter to process the update command - it does not affect any
        /// persistent connection made when creating the class
        /// </summary>
        /// <param name="provider">The provider string for connecting to the data source</param>
        /// <param name="password">The encrypted database password</param>
        /// <param name="deleteQuery">The update query to be used when updating the data source</param>
        public void UpdateData(string provider, string password, string updateQuery)
        {
            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(password, "SkimmedMilk");

            // ### ...and build up the full connection string ###
            provider = provider + decryptedPassword + ";";
            
            try
            {
                OleDbDataAdapter newAdapter = new OleDbDataAdapter();
                OleDbConnection newConnection = new OleDbConnection(provider);
                newConnection.Open();
                newAdapter.UpdateCommand = new OleDbCommand(updateQuery, newConnection);
                newAdapter.UpdateCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new Exception("The data could not be updated for the following reason - " + ex.Message);
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
            foreach (string item in queryList)
            {
                string[] fields = item.Split(';');
                m_localAdapter.SelectCommand.CommandText = fields[0];
                m_localAdapter.Fill(dataSet, fields[1]);
                dataSet.Tables[dataSet.Tables.Count - 1].TableName = fields[1];
            }
            return dataSet;
        }

    }

    /// <summary>
    /// Handles the connecting to and updating of data using OLE
    /// </summary>
    /// <remarks>The MultiTableHandler allows you to update multiple tables within a single dataset</remarks>
    public class MultiTableHandler : New_Wrapper.DataHandler
    {

        #region "Fields"

        OleDbConnection m_connection;
        OleDbDataAdapter m_localAdapter;

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
        public MultiTableHandler(string[,] passedQueries, string passedProvider, string passedPassword)
        {

            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(passedPassword, "SkimmedMilk");

            // ### ...and build up the full connection string ###
            passedProvider = passedProvider + decryptedPassword + ";";

            selectQueries = passedQueries;
            providerString = passedProvider;
            connectToData();
        }

        #endregion

        #region "Public methods"

        /// <summary>
        /// Connects to the data source and populates a data set with the requested tables
        /// </summary>
        /// <param name="passedQueries">The table names and SQL queries that the data tables will be based on</param>
        /// <param name="passedProvider">The provider string for connecting to the data source</param>
        public void connectToData(string[,] passedQueries, string passedProvider)
        {
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
        public void connectToData(string[,] passedQueries, string passedProvider, string passedPassword)
        {
            // ### Decrypt the encrypted password... ###
            string decryptedPassword = StringCipher.Decrypt(passedPassword, "SkimmedMilk");
            selectQueries = passedQueries;
            providerString = passedProvider + decryptedPassword + ";";
            connectToData();
        }

        /// <summary>
        /// Connects to the data source and populates a data set with the requested tables
        /// </summary>
        public void connectToData()
        {
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

        /// <summary>
        /// Updates the specified table(s) in the data set
        /// </summary>
        /// <param name="tableName">The name of the table that is to be updated</param>
        public void UpdateData(string tableName)
        {
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

        /// <summary>
        /// Adds a table from the multihandler instance to an external dataset
        /// </summary>
        /// <param name="passedDataset">The dataset that the table is to be copied to</param>
        /// <param name="passedTableName">The name of the table to be copied to the external dataset</param>
        /// <remarks>The table being copied must already exist in the sourceData property of the class</remarks>
        public void AddTableToDataset(ref DataSet passedDataset, string passedTableName)
        {
            m_localAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;
            //m_localAdapter.Fill(passedDataset, passedTableName);
            DataTable myTable = sourceData.Tables[passedTableName].Clone();
            //sourceData.Tables[passedTableName].Merge();
            passedDataset.Tables.Add(myTable);
            passedDataset.Tables[passedTableName].Merge(sourceData.Tables[passedTableName]);
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

        #endregion

    }

}
