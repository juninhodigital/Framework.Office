using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace Framework.Office
{
    /// <summary>
    /// This class is in charge of providing methods to import CSV and XSL files using the OLEDB providers
    /// </summary>
    public class Excel
    {
        #region| Methods |

        /// <summary>
        /// Initializes a new instance of the System.Data.OleDb.OleDbConnection class with the specified connection string.
        /// </summary>
        /// <param name="FilePath">The file path used to open the datasource</param>
        /// <returns>OleDbConnection</returns>
        private static OleDbConnection GetOleDBConnection(string FilePath)
        {
            if (FilePath.EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase))
            {
                return new OleDbConnection(@"Provider=Microsoft.Jet.Oledb.4.0;Data Source=" + FilePath + ";Extended Properties=Excel 8.0;");
            }
            else
            {
                return new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=Excel 8.0;");
            }
        }

        /// <summary>
        /// Returns schema information from a data source as indicated by a GUID, and after it applies the specified restrictions.
        /// </summary>
        /// <param name="FilePath">FilePath</param>
        /// <returns>System.Data.</returns>
        private static System.Data.DataTable GetSchema(string FilePath)
        {
            var oDataTable = new System.Data.DataTable();

            OleDbConnection oConnection = null;

            try
            {
                oConnection = GetOleDBConnection(FilePath);

                oConnection.Open();

                oDataTable = oConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                oConnection.Close();
                oConnection.Dispose();

                return oDataTable;
            }
            catch (Exception oErro)
            {
                oConnection.Close();
                oConnection.Dispose();

                throw oErro;
            }
        }

        /// <summary>
        /// Returns the Tabs page name that exists in the Excel File
        /// </summary>
        /// <param name="FilePath">The file path used to open the datasource</param>
        /// <returns>List of String</returns>
        public static List<string> GetTabsFromSheet(string FilePath)
        {
            return GetTabs(FilePath).Distinct().ToList();
        }

        /// <summary>
        /// Returns the Tabs page name that exists in the Excel File
        /// </summary>
        /// <param name="FilePath">The file path used to open the datasource</param>
        /// <returns>List of String</returns>
        private static IEnumerable<string> GetTabs(string FilePath)
        {
            var oDataTable = Excel.GetSchema(FilePath);

            if (oDataTable != null && oDataTable.Rows.Count > 0)
            {
                foreach (DataRow oRow in oDataTable.Rows)
                {
                    var TabName = oRow["TABLE_NAME"].ToString().Replace("$", "");
                    yield return TabName;
                }
            }
        }

        /// <summary>
        /// Returns a System.Data.Datatable whose DataSource will be filled from the Excel file using OLEDB
        /// </summary>
        /// <param name="FilePath">FilePath</param>
        /// <param name="SheetName">Sheet Name in the Excel File</param>
        /// <returns>System.Data.DataTable</returns>
        public static System.Data.DataTable ImportUsingOLEDB(string FilePath, string SheetName)
        {
            var oDataTable = new System.Data.DataTable();

            OleDbConnection oConnection = null;
            OleDbCommand oCommand = null;

            var SQL = string.Empty;

            if (SheetName.Contains("$"))
            {
                SQL = "SELECT * FROM [" + SheetName + "]";
            }
            else
            {
                SQL = "SELECT * FROM [" + SheetName + "$]";
            }

            try
            {
                oConnection = GetOleDBConnection(FilePath);
                oCommand = new OleDbCommand(SQL, oConnection);

                oConnection.Open();

                var oReader = oCommand.ExecuteReader();

                if (oReader != null && oReader.HasRows)
                {
                    oDataTable.Load(oReader);

                    oCommand.Dispose();

                    oConnection.Close();
                    oConnection.Dispose();

                    return oDataTable;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception oErro)
            {
                oConnection.Close();
                oConnection.Dispose();

                throw oErro;
            }
        }

        /// <summary>
        /// Returns a System.Data.Datatable whose DataSource will be filled from a CSV file
        /// </summary>
        /// <param name="FilePath">FilePath</param>
        /// <param name="delimiter">Delimiter</param>
        /// <returns>System.Data.DataTable</returns>
        public static DataTable ImportFromCSV(string FilePath, char delimiter = ';')
        {
            var output = new DataTable();

            using (var reader = new StreamReader(File.OpenRead(FilePath)))
            {
                var hasHeader = false;

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(delimiter);

                    if (!hasHeader && values!=null && values.Length>0)
                    {
                        foreach (var value in values)
                        {
                            var column = new DataColumn(value);
                            output.Columns.Add(column);
                        }

                        hasHeader = true;
                    }
                    else
                    {
                        if (values!=null && values.Length > 0)
                        {
                            var row = output.NewRow();

                            for (int i = 0; i < values.Length; i++)
                            {
                                row[i] = values[i];
                            }

                            output.Rows.Add(row);
                        }
                    }
                }
            }

            return output;
        }

        #endregion
    }
}
