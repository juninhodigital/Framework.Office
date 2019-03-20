using System;
using System.Collections.Generic;
using System.Data;

using ClosedXML.Excel;

namespace Framework.Office
{
    /// <summary>
    /// This class is in charge of providing methods to import CSV, XSL and XSLX files using the ClosedXML.Excel
    /// </summary>
    public static class Excel12
    {
        #region| Methods |

        /// <summary>
        /// Returns the Tabs page name that exists in the Excel File
        /// </summary>
        /// <param name="filePath">The file path used to open the datasource</param>
        /// <returns>List of String</returns>
        public static List<string> GetTabsFromSheet(string filePath)
        {
            var output = new List<string>();

            using (var workbook = GetWorkbook(filePath))
            {
                foreach (var item in workbook.Worksheets)
                {
                    output.Add(item.Name);
                }
            }

            return output;
        }

        /// <summary>
        /// Returns a System.Data.Datatable whose DataSource will be filled from the Excel file
        /// </summary>
        /// <param name="filePath">FilePath</param>
        /// <param name="sheetName">Sheet Name in the Excel File</param>
        /// <returns>System.Data.DataTable</returns>
        public static DataTable Import(string filePath, string sheetName)
        {
            DataTable output = null;

            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (var workbook = GetWorkbook(filePath))
            {
                foreach (var workSheet in workbook.Worksheets)
                {
                    if (workSheet.Name.Equals(sheetName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        //Create a new DataTable.
                        output = new DataTable();

                        //Loop through the Worksheet rows.
                        bool firstRow = true;
                        foreach (IXLRow row in workSheet.Rows())
                        {
                            //Use the first row to add columns to DataTable.
                            if (firstRow)
                            {
                                foreach (IXLCell cell in row.Cells())
                                {
                                    output.Columns.Add(cell.Value.ToString());
                                }
                                firstRow = false;
                            }
                            else
                            {
                                //Add rows to DataTable.
                                output.Rows.Add();
                                int i = 0;

                                //foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                                foreach (var cell in row.Cells())
                                {
                                    output.Rows[output.Rows.Count - 1][i] = cell.Value.ToString();
                                    i++;
                                }
                            }
                        }

                        break;
                    }
                }

                return output;
            }
        }

        /// <summary>
        /// Opens an existing workbook from a file.
        /// </summary>
        /// <param name="filePath">file to open</param>
        /// <returns>XLWorkbook</returns>
        public static XLWorkbook GetWorkbook(string filePath)
        {
            var workbook = new XLWorkbook(filePath);

            return workbook;
        }

        #endregion
    }
}
