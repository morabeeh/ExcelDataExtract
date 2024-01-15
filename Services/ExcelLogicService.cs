using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using System.Text;

namespace ExcelDataExtract.Services
{   

    public interface IExcelLogicService
    {
        bool GetExcelData(IFormFile postedFile);
    }
    public class ExcelLogicService:IExcelLogicService
    {
        private readonly ILogger<ExcelLogicService> _logger;
        
        private IHostingEnvironment Environment;
        private IConfiguration _configuration { get; set; }

        #region Constants
        //Read the connection string for the Excel file.
        private static string excelConString;
        private static string dbConString;
        private string sheetName;
        #endregion
        public ExcelLogicService(ILogger<ExcelLogicService> logger, IHostingEnvironment _environment, IConfiguration configuration)
        {
            _logger = logger;
            Environment = _environment;
            _configuration = configuration;
            excelConString= _configuration["ConnectionStrings:ExcelConString"];
            dbConString= _configuration["ConnectionStrings:constr"];
        }

        //to get the excel data
        public bool GetExcelData(IFormFile postedFile)
        {
            try
            {
                _logger.LogError($"Started first method");
                bool success = false;
                if (postedFile != null)
                {
                    //Create a Folder.
                    string path = Path.Combine(this.Environment.WebRootPath, "Uploads");
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    //Save the uploaded Excel file.
                    string fileName = Path.GetFileName(postedFile.FileName);
                    string filePath = Path.Combine(path, fileName);
                    using (FileStream stream = new FileStream(filePath, FileMode.Create))
                    {
                        postedFile.CopyTo(stream);
                    }

                    //Read the connection string for the Excel file.
                    DataTable datatable = new DataTable();
                    excelConString = string.Format(excelConString, filePath);

                    using (OleDbConnection connExcel = new OleDbConnection(excelConString))
                    {
                        using (OleDbCommand cmdExcel = new OleDbCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = connExcel;

                                //Get the name of First Sheet.
                                connExcel.Open();
                                DataTable dtExcelSchema;
                                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                connExcel.Close();

                                //Read Data from First Sheet.
                                connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(datatable);
                                connExcel.Close();
                                success= true;
                            }
                        }
                    }
                    //to check Data extraction from Excel worked successfully.
                    if (success)
                    {
                        CreateDynamicTable(datatable, postedFile);
                    }
                }
                return success;
            }
            catch (Exception ex)
            {
                _logger.LogError($"An error occurred: {ex.Message}");
                return false;

            }
        }


        //dynamically mapping the datatable into dynamically created table in database
        private bool CreateDynamicTable(DataTable datatable, IFormFile postedFile)
        {
            try
            {
                bool success = false;

                using (SqlConnection con = new SqlConnection(dbConString))
                {
                    con.Open();

                    // Set the initial destination table name based on the Excel file name + sheet name
                    string baseDestinationTableName = "dbo." + Path.GetFileNameWithoutExtension(postedFile.FileName) + "_" + sheetName.Replace("$", "");

                    // Check for spaces in the destination table name and replace them with underscores
                    baseDestinationTableName = baseDestinationTableName.Replace(" ", "_");

                    string destinationTableName = baseDestinationTableName;

                    // Check if the destination table already exists
                    int tableSuffix = 1;
                    while (TableExists(con, destinationTableName))
                    {
                        // If the table already exists, append a numeric suffix and check again
                        destinationTableName = $"{baseDestinationTableName}{tableSuffix}";
                        tableSuffix++;
                    }

                    // Create the destination table dynamically
                    StringBuilder createTableSql = new StringBuilder($"CREATE TABLE {destinationTableName} (Id INT PRIMARY KEY IDENTITY(1,1), ");
                    foreach (DataColumn column in datatable.Columns)
                    {
                        createTableSql.Append($"[{column.ColumnName}] NVARCHAR(MAX), ");  // Enclose column name in square brackets
                    }
                    createTableSql.Length -= 2;
                    createTableSql.Append(")");

                    using (SqlCommand cmdCreateTable = new SqlCommand(createTableSql.ToString(), con))
                    {
                        cmdCreateTable.ExecuteNonQuery();
                        success = true;
                    }

                    if (success)
                    {
                        // Use destinationTableName for the subsequent operations
                        using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                        {
                            sqlBulkCopy.DestinationTableName = destinationTableName;

                            // Dynamically generate ColumnMappings based on DataTable columns
                            foreach (DataColumn column in datatable.Columns)
                            {
                                sqlBulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                            }

                            // Perform the bulk insert
                            sqlBulkCopy.WriteToServer(datatable);
                        }
                    }

                    con.Close();
                }
                return success;
            }
            catch (Exception ex)
            {
                _logger.LogError($"An error occurred in CreateDynamicTable: {ex.Message}");
                return false;
            }
        }
        // Function to check if a table exists in the database
        private bool TableExists(SqlConnection connection, string tableName)
        {
            try
            {
                using (SqlCommand command = new SqlCommand($"IF OBJECT_ID('{tableName}', 'U') IS NOT NULL SELECT 1 ELSE SELECT 0", connection))
                {
                    return (int)command.ExecuteScalar() == 1;
                }
            }
            catch(Exception ex)
            {
                _logger.LogError($"An error occurred in checking Table Exists checking: {ex.Message}");
                return false;
            }

        }

    }
}
