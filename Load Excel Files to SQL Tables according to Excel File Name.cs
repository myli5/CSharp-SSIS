#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;

#endregion

namespace ST_c20de9a050d848109c74ad14700a4ca5
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
	public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
	{
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion


		/// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
		public void Main()
		{
			// TODO: Add your code here
            String FolderPath=Dts.Variables["User::FolderPath"].Value.ToString();
            String TableName = "";
            String SchemaName = Dts.Variables["User::SchemaName"].Value.ToString();
            var directory = new DirectoryInfo(FolderPath);
            FileInfo[] files = directory.GetFiles();
            
            //Declare and initilize variables
            string fileFullPath = "";
            
            //Get one Book(Excel file at a time)
            foreach (FileInfo file in files)
            {
                   fileFullPath = FolderPath+"\\"+file.Name;
                                   
                //Create Excel Connection
                string ConStr;
                string HDR;
                HDR="YES";
                ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileFullPath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                    OleDbConnection cnn = new OleDbConnection(ConStr);
            
                //Remove All Numbers and other characters and leave alphabets for name
                    System.Text.RegularExpressions.Regex rgx = new System.Text.RegularExpressions.Regex("[^a-zA-Z]");
                    TableName = rgx.Replace(file.Name, "").Replace("xlsx","");
                //MessageBox.Show(TableName);
                //Get Sheet Name
                   cnn.Open();
                DataTable dtSheet = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetname;
                sheetname="";
           foreach (DataRow drSheet in dtSheet.Rows)
            {
                if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                {
                     sheetname=drSheet["TABLE_NAME"].ToString();
                     
                     //Load the DataTable with Sheet Data so we can get the column header
                     OleDbCommand oconn = new OleDbCommand("select top 1 * from [" + sheetname + "]", cnn);
                     OleDbDataAdapter adp = new OleDbDataAdapter(oconn);
                     DataTable dt = new DataTable();
                     adp.Fill(dt);
                     cnn.Close();

                     //Prepare Header columns list so we can run against Database to get matching columns for a table.
                     string ExcelHeaderColumn = "";
                     string SQLQueryToGetMatchingColumn = "";
                     for (int i = 0; i < dt.Columns.Count; i++)
                     {
                         if (i != dt.Columns.Count - 1)
                             ExcelHeaderColumn += "'" + dt.Columns[i].ColumnName + "'" + ",";
                         else
                             ExcelHeaderColumn += "'" + dt.Columns[i].ColumnName + "'";
                     }

                     SQLQueryToGetMatchingColumn = "select STUFF((Select  ',['+Column_Name+']' from Information_schema.Columns where Table_Name='" +
                         TableName + "' and Table_SChema='" + SchemaName + "'" +
                         "and Column_Name in (" + @ExcelHeaderColumn + ") for xml path('')),1,1,'') AS ColumnList";

                     // MessageBox.Show(SQLQueryToGetMatchingColumn);
                     //MessageBox.Show(ExcelHeaderColumn);

                     //USE ADO.NET Connection
                     SqlConnection myADONETConnection = new SqlConnection();
                     myADONETConnection = (SqlConnection)(Dts.Connections["DBConn"].AcquireConnection(Dts.Transaction) as SqlConnection);

                     //Get Matching Column List from SQL Server
                     string SQLColumnList = "";
                     SqlCommand cmd = myADONETConnection.CreateCommand();
                     cmd.CommandText = SQLQueryToGetMatchingColumn;
                     SQLColumnList = (string)cmd.ExecuteScalar();

                     //MessageBox.Show(" Matching Columns: " + SQLColumnList);


                     //Use Actual Matching Columns to get data from Excel Sheet
                     OleDbConnection cnn1 = new OleDbConnection(ConStr);
                     cnn1.Open();
                     OleDbCommand oconn1 = new OleDbCommand("select " + SQLColumnList + " from [" + sheetname + "]", cnn1);
                     OleDbDataAdapter adp1 = new OleDbDataAdapter(oconn1);
                     DataTable dt1 = new DataTable();
                     adp1.Fill(dt1);
                     cnn1.Close();


                     //Load Data from DataTable to SQL Server Table.
                     using (SqlBulkCopy BC = new SqlBulkCopy(myADONETConnection))
                     {
                         BC.DestinationTableName = SchemaName + "." + TableName;
                         foreach (var column in dt1.Columns)
                             BC.ColumnMappings.Add(column.ToString(), column.ToString());
                         BC.WriteToServer(dt1);
                     }

                }
                } 
            }
           
			Dts.TaskResult = (int)ScriptResults.Success;

		}

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

	}

}