using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;

namespace ReferenceScanner
{

    public enum FileType { xls, xlsx, xlsm, csv};
    public enum OutputFormat { Excel, CSV, DataTable};
    
    public class ReferenceCheckedData
    {

#region Private Methods

        ///// <summary>
        ///// Operates on every row and column in the supplied data source.
        ///// </summary>
        ///// <param name="filePath">FilePath to where the excel or csv file is located.</param>
        ///// <param name="outputFormat">What data structure the output should be.</param>
        ///// <param name="inputExtention">File extension of the input file.</param>
        ///// <param name="sheetName">Optional Param for the sheet name if input is of type Excel</param>
        ///// <returns></returns>
        //private DataTable OutputScannedData(string filePath, OutputFormat outputFormat, FileType inputExtention, string sheetName = "Sheet1")
        //{
        //    string strFileName = filePath;
        //    New_Wrapper.DataHandler myHandler = new New_Wrapper.DataHandler();

        //    string fileLocation = Path.GetDirectoryName(strFileName);

        //    DataTable cleanedData = new DataTable();

        //    if (inputExtention != FileType.csv)
        //    {
        //        cleanedData = myHandler.ReturnExcelSheetAsDataTable(strFileName, sheetName, true);
        //    }
        //    else
        //    {
        //        cleanedData = myHandler.ReturnCSVAsDataTable(strFileName, true);
        //    }

        //    if (cleanedData.Columns.Count > 11)
        //    {
        //        int i = 0;
        //        for (int j = cleanedData.Columns.Count - 1; j >= 10; j--)
        //        {
        //            if (cleanedData.Columns[j].ColumnName.Contains(Path.GetFileNameWithoutExtension(strFileName)))
        //            {
        //                cleanedData.Columns.RemoveAt(j);
        //            }
        //            i++;
        //        }
        //    }

        //    cleanedData.Columns.Add("NewReference", typeof(string));
        //    cleanedData.Columns.Add("HOD", typeof(string));

        //    ReferenceParser.ReferenceParser rp = new ReferenceParser.ReferenceParser();

        //    foreach (DataRow dr in cleanedData.Rows)
        //    {
        //        bool found = false;
        //        foreach (DataColumn dc in cleanedData.Columns)
        //        {

        //            if (found == false)
        //            {
        //                dr["NewReference"] = "";
        //                dr["HOD"] = "";
        //                foreach (string s in rp.Parse(dr[dc.ColumnName].ToString(), true, true))
        //                {

        //                    if (s.Trim() != "")
        //                    {
        //                        string[] splitRefHod = s.Split(new string[] { "??$??" }, StringSplitOptions.None);
        //                        dr["NewReference"] = splitRefHod[0];
        //                        dr["HOD"] = splitRefHod[1];
        //                        found = true;
        //                        break;
        //                    }
        //                    else
        //                    {
        //                        found = false;
        //                        dr["NewReference"] = "";
        //                        dr["HOD"] = "";
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return cleanedData;
        //}


        /// <summary>
        /// Overload - Operatate on defined columns provided by a list.
        /// </summary>
        /// <param name="filePath">FilePath to where the excel or csv file is located.</param>
        /// <param name="outputFormat">What data structure the output should be.</param>
        /// <param name="inputExtention">File extension of the input file.</param>
        /// <param name="columnHeaders">List of columns headers to parse for references.</param>
        /// <param name="sheetName">Optional Param for the sheet name if input is of type Excel</param>
        /// <returns>Datatable</returns>
        private DataTable OutputScannedData(string filePath, OutputFormat outputFormat, FileType inputExtention, List<string> columnHeaders, string sheetName = "Sheet1")
        {
            string strFileName = filePath;
            New_Wrapper.DataHandler myHandler = new New_Wrapper.DataHandler();

            string fileLocation = Path.GetDirectoryName(strFileName);

            DataTable cleanedData = new DataTable();

            if (inputExtention != FileType.csv)
            {
                cleanedData = myHandler.ReturnExcelSheetAsDataTable(strFileName, sheetName, true);
            }
            else
            {
                cleanedData = myHandler.ReturnCSVAsDataTable(strFileName, true);
            }

            if (cleanedData.Columns.Count > 11)
            {
                int i = 0;
                for (int j = cleanedData.Columns.Count - 1; j >= 10; j--)
                {
                    if (cleanedData.Columns[j].ColumnName.Contains(Path.GetFileNameWithoutExtension(strFileName)))
                    {
                        cleanedData.Columns.RemoveAt(j);
                    }
                    i++;
                }
            }

            cleanedData.Columns.Add("NewReference", typeof(string));
            cleanedData.Columns.Add("HOD", typeof(string));

            ReferenceParser.ReferenceParser rp = new ReferenceParser.ReferenceParser();

            foreach (DataRow dr in cleanedData.Rows)
            {
                bool found = false;
                foreach (DataColumn dc in cleanedData.Columns)
                {
                    if (columnHeaders.Contains(dc.ColumnName))
                    {
                        if (found == false)
                        {
                            dr["NewReference"] = "";
                            dr["HOD"] = "";
                            foreach (string s in rp.Parse(dr[dc.ColumnName].ToString(), true, true))
                            {

                                if (s.Trim() != "")
                                { 
                                    string[] splitRefHod = s.Split(new string[] { "??$??" }, StringSplitOptions.None);
                                    dr["NewReference"] = splitRefHod[0];
                                    dr["HOD"] = splitRefHod[1];
                                    found = true;
                                    break;
                                }
                                else
                                {
                                    found = false;
                                    dr["NewReference"] = "";
                                    dr["HOD"] = "";
                                }
                            }
                    
                        }
                    }
                }
            }
            return cleanedData;
        }

        /// <summary>
        /// 
        /// </summary>
        private List<string> columnHeaders;
    
        /// <summary>
        /// Output a CSV file from the DataTable returned from the OutputScannedData Method.
        /// </summary>
        /// <param name="outputLocation">Where the file should be output too.</param>
        /// <param name="fileName">new name for the output file.</param>
        /// <param name="cleanedData">provide a datatable.</param>
        private void ReturnReformatDataTableAsCsv(DataTable cleanedData, string outputLocation, string fileName)
        {
            New_Wrapper.DataHandler csvHandler = new New_Wrapper.DataHandler();
            csvHandler.GenerateCSV(cleanedData, outputLocation, fileName, "#");
        }

        /// <summary>
        /// Output a Excel workbook for the Datatable returned from OutputScannedata Method.
        /// </summary>
        /// <param name="outputLocation">Where the file should be output too.</param>
        /// <param name="fileName">new name for the output file.</param>
        /// <param name="cleanedData">provide a datatable.</param>
        /// <param name="sheetname">The sheet where the data will be exporeted from</param>
        private void ReturnReformatDataTableAsExcel(DataTable cleanedData, string outputLocation, string fileName, string sheetname = "Sheet1")
        {
            cleanedData.ExportToExcel(fileName, outputLocation);
        }

#endregion

#region public methods

        public List<string> ColumnHeaders
        {
            get { return columnHeaders; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileLocation">Where the file should be output too.</param>
        /// <param name="fileName">new name for the output file.</param>
        /// <param name="outputFormat">What data structure the output should be.</param>
        /// <param name="fileLocation">Where the file should be output too.</param>
        /// <param name="sheetName">Optional for Excel: The sheet where the data will be exporeted from</param>
        public void ReturnReferenceScannedData(string fileLocation, OutputFormat outputFormat, List<string> columnHeaders, string outputLocation, string fileName, string sheetName = "Sheet1")
        {
        
            switch(outputFormat)
            {
                case OutputFormat.CSV:
                    ReturnReformatDataTableAsCsv(OutputScannedData(fileLocation, outputFormat, FileType.csv, columnHeaders), outputLocation, fileName);
                    break;
                case OutputFormat.Excel:
                    ReturnReformatDataTableAsExcel(OutputScannedData(fileLocation, outputFormat, FileType.xls, columnHeaders, sheetName), outputLocation, fileName);
                    break;
                case OutputFormat.DataTable:
                    throw new Exception("Use the data table overload ya bampot");
                default:
                    break;
            }
        }

        /// <summary>
        /// Returns a datatable of Data scanned for references.
        /// </summary>
        /// <param name="filePath">FilePath to where the excel or csv file is located.</param>
        /// <param name="outputFormat">What data structure the output should be.</param>
        /// <param name="inputExtention">File extension of the input file.</param>
        /// <param name="sheetName">Optional Param for the sheet name if input is of type Excel</param>
        /// <returns>Returns a datatable</returns>
        public DataTable ReturnReferenceScannedDataTable(string filePath, OutputFormat outputFormat, FileType inputExtension, List<string> columnHeaders, string sheetName = "Sheet1")
        {
            switch(inputExtension)
	        {   
                case FileType.csv:
                    return OutputScannedData(filePath, OutputFormat.DataTable, FileType.csv, columnHeaders);
                case FileType.xls:
                    return OutputScannedData(filePath, OutputFormat.DataTable, FileType.xls, columnHeaders, sheetName);
                case FileType.xlsm:
                    return OutputScannedData(filePath, OutputFormat.DataTable, FileType.xlsm, columnHeaders, sheetName);
                case FileType.xlsx:
                    return OutputScannedData(filePath, OutputFormat.DataTable, FileType.xlsx, columnHeaders, sheetName);
		        default:
                    return null;
	        }
        }

        

#endregion 

    }
}
