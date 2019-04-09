using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.OleDb;
using System.IO;


namespace ReferenceScanner
{
    public static class My_DataTable_Extensions
    {
        /// <summary>
        /// This most excellent chap takes a datatable and dumps out an Excel spreadsheet.
        /// 
        /// Extension to the datatable method.
        /// </summary>
        /// <param name="DataTable">Provide a DataTable</param>
        /// <param name="excelFilePath">Provide a filepath to file.</param>
        public static void ExportToExcel(this System.Data.DataTable DataTable, string fileName, string excelFilePath = null)
        {
            try
            {
                int ColumnsCount;

                if (DataTable == null || (ColumnsCount = DataTable.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // loading excel and creating a new workbook
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Excel.Workbooks.Add();

                // creat a single worksheet
                Microsoft.Office.Interop.Excel._Worksheet Worksheet = Excel.ActiveSheet;

                object[] Header = new object[ColumnsCount];

                // loop through column headings
                for (int i = 0; i < ColumnsCount; i++)
                {
                    Header[i] = DataTable.Columns[i].ColumnName;
                }


                Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)
                    (Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
                HeaderRange.Value = Header;
                HeaderRange.Font.Bold = true;

                // DataCells
                int RowsCount = DataTable.Rows.Count;
                object[,] Cells = new object[RowsCount, ColumnsCount];

                for (int j = 0; j < RowsCount; j++)
                {
                    for (int i = 0; i < ColumnsCount; i++)
                    {
                        Cells[j, i] = DataTable.Rows[j][i];
                    }
                }

                Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]),
                    (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount + 1])).Value = Cells;

                // for file path locations
                if (excelFilePath != null && excelFilePath != "")
                {
                    try
                    {
                        Worksheet.SaveAs(excelFilePath + "cleaned_" + fileName);
                        Excel.Quit();
                    }
                    catch (Exception ex)
                    {

                        throw new Exception("ExportToExcel: Excel file could not be saved! check file path.\n"
                            + ex.Message);
                    }
                }
                else
                {
                    Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {

                throw new Exception("ExportToExcel : \n" + ex.Message);
            }
        }
    }
}
