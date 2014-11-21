using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data;
using System.Data.OleDb;
using System.IO;
using ClosedXML.Excel;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;

namespace OrderTracker
{
    public class ImportExport
    {
        /// <summary>
        /// Converts a Excel file to a Data Table
        /// </summary>
        /// <param name="filePath">String containing path to the excel file</param>
        /// <returns>DataTable containg the imported data</returns>
        public static DataSet ExcelToDataSet(string filePath, bool hasHeaders = true)
        {
            try
            {
                DataSet ds = new DataSet();
                DataTable dtexcel = new DataTable();
                string HDR = hasHeaders ? "Yes" : "No";
                string strConn;
                if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
                else
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";

                OleDbConnection connection = new OleDbConnection(strConn);
                String sheetName = String.Empty;
                connection.Open();
                DataTable oDt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (oDt == null)
                    return null;
                else if (oDt.Rows.Count < 1)
                    return null;
                else
                {
                    foreach (DataRow dRow in oDt.Rows)
                    {
                        if (!dRow["TABLE_NAME"].ToString().Contains("FilterDatabase"))
                        {
                            sheetName = dRow["TABLE_NAME"].ToString();
                            if (sheetName.EndsWith("_")) continue;

                            OleDbCommand command = new OleDbCommand("SELECT * FROM [" + sheetName + "]", connection);
                            OleDbDataAdapter adapter = new OleDbDataAdapter();
                            dtexcel.TableName = sheetName;
                            adapter.SelectCommand = command;
                            adapter.Fill(dtexcel);
                            ds.Tables.Add(dtexcel);
                        }
                    }
                }
                connection.Close();
                //dtexcel = ds.Tables[0];
                return ds;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Converts a CSV File to DataTable
        /// </summary>
        /// <param name="filePath">String containing path to the CSV file</param>
        /// <returns>DataTable containg the imported data</returns>
        public static DataTable CSVToDataTable(string filePath, bool hasHeaders = true)
        {
            try
            {
                DataTable csvTable = new DataTable();
                StreamReader reader = new StreamReader(filePath);
                //FileStream reader = new FileStream(filePath, FileMode.Open);
                string[] columnNames;
                columnNames = reader.ReadLine().Split(',');
                for (int index = 0; index < columnNames.Length; index++)
                {
                    if (!hasHeaders)
                        csvTable.Columns.Add("Column" + index);
                    else
                        csvTable.Columns.Add(columnNames[index]);
                }
                if (!hasHeaders)
                {
                    // close and reopen
                    reader.Close();
                    reader = new StreamReader(filePath);
                }
                String[] rowContent;
                while (!reader.EndOfStream)
                {
                    rowContent = reader.ReadLine().Split(',');
                    DataRow dRow = csvTable.NewRow();
                    for (int colCount = 0; colCount < csvTable.Columns.Count; colCount++)
                    {
                        dRow[colCount] = rowContent[colCount];
                    }
                    csvTable.Rows.Add(dRow);
                }

                reader.Close();
                reader.Dispose();
                return csvTable;
            }
            catch
            {
                return null;
            }
        }

        public static DataTable ExcelToDataSetCloseXml(string filePath)
        {
            var datatable = new DataTable();
            // var dataset = new DataSet();
            var workbook = new XLWorkbook(filePath);
            var xlWorksheet = workbook.Worksheet(1);

            //IXLRange xlRangeRow = xlWorksheet.AsRange();                        
            var range = xlWorksheet.Range(xlWorksheet.FirstCellUsed(), xlWorksheet.LastCellUsed());
            // IXLCell rowCell = xlWorksheet.LastCellUsed();

            int col = range.ColumnCount();
            int row = range.RowCount();

            // add columns hedars
            datatable.Clear();

            for (int i = 1; i <= col; i++)
            {
                IXLCell column = xlWorksheet.Cell(1, i);
                datatable.Columns.Add(column.Value.ToString());
            }

            // add rows data   
            int firstHeadRow = 0;
            foreach (var item in range.Rows())
            {
                if (firstHeadRow != 0)
                {
                    var array = new object[col];
                    for (int y = 1; y <= col; y++)
                    {
                        array[y - 1] = item.Cell(y).Value;
                    }
                    datatable.Rows.Add(array);
                }
                firstHeadRow++;
            }
            return datatable;
        }

        public static String PDFText(String PDFFilePath)
        {
            PDDocument doc = PDDocument.load(PDFFilePath);
            PDFTextStripper stripper = new PDFTextStripper();
            return stripper.getText(doc);
        }

    }
}
