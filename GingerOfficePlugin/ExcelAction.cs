#region License
/*
Copyright © 2014-2018 European Support Limited

Licensed under the Apache License, Version 2.0 (the "License")
you may not use this file except in compliance with the License.
You may obtain a copy of the License at 

http://www.apache.org/licenses/LICENSE-2.0 

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
See the License for the specific language governing permissions and 
limitations under the License. 
*/
#endregion

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GingerPlugInsNET.ActionsLib;
using GingerPlugInsNET.PlugInsLib;
using GingerPlugInsNET.ServicesLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace StandAloneActions
{
    public class ExcelAction : PluginServiceBase, IStandAloneAction
    {
        public override string Name { get { return "ExcelService"; } }

        #region Actions

        /// <summary>
        /// Get first row where column Name equal value
        /// </summary>
        /// <param name="GA"></param>
        /// <param name="FileName"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        [GingerAction("ReadExcelCell", "Read From Excel")]
        public void ReadExcelCell(ref GingerAction GA, string FileName, string sheetName, string row, string column)
        {
            try
            {
                if (column.Contains("#") && row.Contains("#"))              //row="#3", column="#B"
                {
                    string col = GetColumnName(FileName, sheetName, column);
                    string nRow = Convert.ToString(GetRowIndex(row));
                    string txt1 = GetCellValue(FileName, "Sheet1", col + nRow);

                    GA.Output.Add("Value", txt1);
                    GA.ExInfo = "Read FileName: " + FileName + " row= " + row + " col= " + column;
                }
                else                                                        //row="First='Moshe'", column="ID"
                {
                    List<string> colList = GetCoulmnsName(FileName, sheetName, column);
                    ReadExcelRowWithCondition(GA, FileName, sheetName, row, colList);
                }
            }
            catch (Exception ex)
            {
                GA.AddError("ReadExcelCell", ex.StackTrace);
            }
        }

        /// <summary>
        /// This action will read the row details for the currentrow index
        /// </summary>
        /// <param name="GA"></param>
        /// <param name="FileName"></param>
        /// <param name="row"></param>
        /// <param name="columns"></param>
        [GingerAction("ReadExcelRow", "Read Next Row From Excel")]
        public void ReadExcelRow(ref GingerAction GA, string FileName, string sheetName, string row, string columns)
        {
            try
            {
                if (!string.IsNullOrEmpty(row) && row.Contains("#"))        //row = "#3"
                {
                    List<string> colList = GetCoulmnsName(FileName, sheetName, columns);
                    int rowNum = GetRowIndex(row);
                    List<string> rowValues = GetCurrentRowValues(FileName, "Sheet1", rowNum, colList);
                    foreach (var item in rowValues)
                    {
                        GA.Output.Add("Value", item.Trim()); 
                    }
                    GA.ExInfo = "Read FileName: " + FileName + " row= " + row;
                }
                else                                                        //row = "Used='No'"
                {                                                           //row = "ID>'30' and Used='No'"
                    List<string> colList = GetCoulmnsName(FileName, sheetName, columns);
                    ReadExcelRowWithCondition(GA, FileName, sheetName, row, colList);
                }
            }
            catch (Exception ex)
            {
                GA.AddError("ReadExcelRow", ex.StackTrace);
            }
        }

        /// <summary>
        /// Read and update cell value with updatevalue where the currentvalue value for the columnName provided
        /// </summary>
        /// <param name="GA"></param>
        /// <param name="FileName"></param>
        /// <param name="row"></param>
        /// <param name="columns"></param>
        /// <param name="values"></param>
        [GingerAction("ReadExcelAndUpdate", "Read and Update Excel Cell")]
        public void ReadExcelAndUpdate(ref GingerAction GA, string FileName, string sheetName, string row, string columns, string values)
        {
            try
            {
                if (!string.IsNullOrEmpty(row) && !string.IsNullOrEmpty(values))
                {                    
                    int rowNum = GetRowIndex(row);
                    bool isUpdated = UpdateRowCells(FileName, sheetName, rowNum, values);
                    if (!isUpdated)
                    {
                        GA.AddError("ReadExcelAndUpdate", "Update Fail!");
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(columns))
                        {
                            GA.Output.Add("Value", Convert.ToString(isUpdated));
                        }
                        else
                        {
                            ReadExcelRow(ref GA, FileName, sheetName, row, columns);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                GA.AddError("ReadExcelAndUpdate", ex.StackTrace);
            }
        }

        /// <summary>
        /// This action is used to append data to the excel
        /// </summary>
        /// <param name="GA"></param>
        /// <param name="FileName"></param>
        /// <param name="rowValueList"></param>
        [GingerAction("AppendData", "Append Data to Excel")]
        public void AppendData(ref GingerAction GA, string FileName, string sheetName, string values)
        {
            try
            {
                // Appends new row in sheet
                List<ExcelCellValues> rowValueList = GetColumnUpdateValues(FileName, sheetName, values);
                var newRow = AppendRowExcel(FileName, sheetName, rowValueList);
                uint rowCount = newRow.RowIndex;

                GA.Output.Add("Value", Convert.ToString(rowCount));

                GA.ExInfo = "Read FileName: " + FileName;
            }
            catch (Exception ex)
            {
                GA.AddError("AppendData", ex.StackTrace);
            }
        }
        
        /// <summary>
        /// This action is used to add the value to cell
        /// </summary>
        /// <param name="GA"></param>
        /// <param name="FileName"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        [GingerAction("WriteExcel", "Write to Excel")]
        public void WriteExcel(ref GingerAction GA, string FileName, string sheetName, int row, string column, string value)
        {
            //// Create new sheet and insert the value in A1
            //// just as smaple for writing
            bool isInserted = InsertText(FileName, sheetName, (uint)row, column, value);
            if(isInserted)
            {
                GA.Output.Add("Value", Convert.ToString(true));
            }
            else
            {
                GA.AddError("WriteExcel", "Failed to Update");
            }
            GA.ExInfo = "Read FileName: " + FileName + " row= " + row + " col= " + column;
        }
        
        #endregion

        #region Private Methods

        /// <summary>
        /// This method will read the excel with the condition provided
        /// </summary>
        /// <param name="GA"></param>
        /// <param name="FileName"></param>
        /// <param name="row"></param>
        /// <param name="columns"></param>
        private List<string> ReadExcelRowWithCondition(GingerAction GA, string FileName, string sheetName, string row, List<string> columnsList)
        {
            List<string> values = new List<string>();
            try
            {
                values = GetExcelRowWithCondition(FileName, sheetName, row, columnsList);
                foreach (var item in values)
                {
                    GA.Output.Add("Value", Convert.ToString(item.Trim())); 
                }
                GA.ExInfo = "Read FileName: " + FileName + " row= " + row;
            }
            catch (Exception ex)
            {
                throw;
            }
            return values;
        }

        /// <summary>
        /// This method will return the rowIndex
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private int GetRowIndex(string row)
        {
            int rIndex = -1;
            try
            {
                if (row.StartsWith("#"))
                {
                    string nRow = row.Replace("#", "");
                    int.TryParse(nRow, out rIndex);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return rIndex;
        }
        
        /// <summary>
        /// This method gets the cell value
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <param name="addressName"></param>
        /// <returns></returns>
        private string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;
            try
            {
                // Open the spreadsheet document for read-only access.
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    // Retrieve a reference to the workbook part.
                    WorkbookPart wbPart = document.WorkbookPart;

                    // Find the sheet with the supplied name, and then use that 
                    // Sheet object to retrieve a reference to the first worksheet.
                    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                    // Throw an exception if there is no sheet.
                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
                    }

                    // Retrieve a reference to the worksheet part.
                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                    // Use its Worksheet property to get a reference to the cell 
                    // whose address matches the address you supplied.
                    Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();

                    // If the cell does not exist, return an empty string.
                    if (theCell != null)
                    {
                        value = theCell.InnerText;

                        // If the cell represents an integer number, you are done. 
                        // For dates, this code returns the serialized value that 
                        // represents the date. The code handles strings and 
                        // Booleans individually. For shared strings, the code 
                        // looks up the corresponding value in the shared string 
                        // table. For Booleans, the code converts the value into 
                        // the words TRUE or FALSE.
                        if (theCell.DataType != null)
                        {
                            switch (theCell.DataType.Value)
                            {
                                case CellValues.SharedString:

                                    // For shared strings, look up the value in the
                                    // shared strings table.
                                    var stringTable =
                                        wbPart.GetPartsOfType<SharedStringTablePart>()
                                        .FirstOrDefault();

                                    // If the shared string table is missing, something 
                                    // is wrong. Return the index that is in
                                    // the cell. Otherwise, look up the correct text in 
                                    // the table.
                                    if (stringTable != null)
                                    {
                                        value =
                                            stringTable.SharedStringTable
                                            .ElementAt(int.Parse(value)).InnerText;
                                    }
                                    break;

                                case CellValues.Boolean:
                                    switch (value)
                                    {
                                        case "0":
                                            value = "FALSE";
                                            break;
                                        default:
                                            value = "TRUE";
                                            break;
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return value;
        }

        /// <summary>
        /// This method will return the row with the complex condition
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="condition"></param>
        /// <returns></returns>
        private List<string> GetExcelRowWithCondition(string fileName, string sheetName, string condition, List<string> columnsList)
        {
            int rowIndex = 0;
            int colIndex = 0;
            List<string> list = new List<string>();
            try
            {
                // Open the spreadsheet document for read-only access.
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    if (rows != null)
                    {
                        var tempRow = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault();
                        var tempcols = tempRow.Elements<Cell>();
                        var colsCount = tempRow.Elements<Cell>().Count();
                        
                        DataTable dt = new DataTable();
                        for (int i = 0; i < colsCount; i++)
                        {
                            string val = GetCurrentCellValue(document, tempcols.ElementAt(i));
                            dt.Columns.Add(val);
                        }
                                                
                        foreach (var row in rows)
                        {
                            if (rowIndex > 0)
                            {
                                colIndex = 0;
                                DataRow dr = dt.NewRow();
                                var cols = row.Elements<Cell>();
                                foreach (var col in cols)
                                {
                                    string cellValue = GetCurrentCellValue(document, col);
                                    dr[colIndex] = cellValue;
                                    colIndex++;
                                }
                                dt.Rows.Add(dr);
                            }
                            rowIndex++;
                        }
                        DataView dataView = dt.DefaultView;
                        dataView.RowFilter = condition;

                        foreach (DataRow dr in dataView.ToTable().Rows)
                        {
                            if (columnsList != null && columnsList.Count > 0)
                            {
                                foreach (var colName in columnsList)
                                {
                                    int indx = ColumnNumber(colName);
                                    list.Add(Convert.ToString(dr[indx-1]).Trim());
                                }
                                break;
                            }
                            else
                            {
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    list.Add(Convert.ToString(dr[i]).Trim());                                    
                                }
                                break;
                            }                            
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return list;
        }
        
        /// <summary>
        /// This method will read all the cell values from the next row of the currRow
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <param name="currRowNo"></param>
        /// <returns></returns>
        public List<string> GetCurrentRowValues(string fileName, string sheetName, int currRowNo, List<string> columnsList)
        {
            List<string> values = new List<string>();
            try
            {
                // Open the spreadsheet document for read-only access.
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    string relationshipId = sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                    var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                    if (columnsList != null && columnsList.Count > 0)
                    {
                        foreach (Cell cell in rows.ElementAt(currRowNo - 1))
                        {
                            foreach (var colName in columnsList)
                            {
                                if (GetExcelColumnIndexFromColumnName(colName) > 0 && cell.CellReference.InnerText.Contains(colName))
                                {
                                    values.Add(GetCurrentCellValue(document, cell));
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (Cell cell in rows.ElementAt(currRowNo - 1))
                        {
                            values.Add(GetCurrentCellValue(document, cell));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return values;
        }

        /// <summary>
        /// This method is used to get the value fro mthe cell
        /// </summary>
        /// <param name="document"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        private string GetCurrentCellValue(SpreadsheetDocument document, Cell cell)
        {
            string value = string.Empty;
            try
            {
                SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                value = cell.CellValue.InnerXml;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return value;
        }
        
        /// <summary>
        /// This method gets the corrected values for cellvalues to update
        /// </summary>
        /// <param name="cellValues"></param>
        /// <returns></returns>
        private List<ExcelCellValues> GetColumnUpdateValues(string fileName, string sheetName, string cellValues)
        {
            List<ExcelCellValues> listCols = new List<ExcelCellValues>();
            try
            {
                string[] values = cellValues.Split(',');
                if (values != null && values.Length > 0)
                {
                    int cIndex = 65; //A
                    foreach (var item in values)
                    {
                        if (item.Contains("="))
                        {
                            string[] colVals = item.Split('=');
                            ExcelCellValues obj = new ExcelCellValues()
                            {
                                CellName = GetColumnName(fileName, sheetName, colVals[0]),
                                CellUpdateValue = colVals[1]
                            };
                            listCols.Add(obj); 
                        }
                        else
                        {
                            ExcelCellValues obj = new ExcelCellValues()
                            {
                                CellName = char.ConvertFromUtf32(cIndex),
                                CellUpdateValue = item
                            };
                            listCols.Add(obj);
                            cIndex++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return listCols;
        }

        #region Columns Operation

        /// <summary>
        /// This method is used to Convert columnName to Number
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private bool IsColumnNameSame(string columnName, string cellColumnName)
        {
            bool isSame = false;
            try
            {
                columnName = columnName.Replace("#", "");
                string cName = string.Empty;
                foreach (char ch in cellColumnName)
                {
                    int num = 0;
                    if (int.TryParse(Convert.ToString(ch), out num))
                    {
                        break;
                    }
                    else
                    {
                        cName = cName + Convert.ToString(ch);
                    }
                }

                if (columnName.ToLower() == cName.ToLower())
                {
                    isSame = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isSame;
        }

        /// <summary>
        /// This method returns the index of column with the coladdress
        /// </summary>
        /// <param name="colAdress"></param>
        /// <returns></returns>
        private int ColumnNumber(string colAdress)
        {
            int[] digits = new int[colAdress.Length];
            for (int i = 0; i < colAdress.Length; ++i)
            {
                digits[i] = Convert.ToInt32(colAdress[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }

        /// <summary>
        /// This method will Get ColumnName From ColumnNumber
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        private string GetColumnNameFromColumnNumber(int columnNumber)
        {
            string setColumnName = String.Empty;
            try
            {
                if (columnNumber > 0)
                {

                    int tempRemainder = 0;
                    while (columnNumber > 0)
                    {
                        tempRemainder = (columnNumber - 1) % 26;
                        setColumnName = Convert.ToChar(65 + tempRemainder).ToString() + setColumnName;
                        columnNumber = (int)((columnNumber - tempRemainder) / 26);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return setColumnName;
        }

        /// <summary>
        /// This method is used to get the columnName
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private int GetExcelColumnIndexFromColumnName(string columnName)
        {
            int validColumn = 0;
            try
            {
                if (!string.IsNullOrEmpty(columnName))
                {
                    columnName = columnName.ToUpperInvariant();
                    for (int i = 0; i < columnName.Length; i++)
                    {
                        validColumn = validColumn * 26;
                        validColumn = validColumn + (columnName[i] - 'A' + 1);
                    }

                    if (validColumn >= 1 && validColumn <= 16384)
                    {
                        validColumn = 1;
                    }
                    else
                    {
                        validColumn = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return validColumn;
        }

        /// <summary>
        /// This method will return the list of columnname
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        private List<string> GetCoulmnsName(string fileName, string sheetName, string columns)
        {
            List<string> list = new List<string>();
            try
            {
                if (!string.IsNullOrEmpty(columns))
                {
                    string[] cols = columns.Split(',');
                    if (cols != null && cols.Length > 0)
                    {
                        foreach (var str in cols)
                        {
                            string colName = string.Empty;
                            if (str.TrimStart().StartsWith("#"))
                            {
                                colName = str.TrimStart().Replace("#", "");
                                int cInd = 0;
                                if (int.TryParse(colName, out cInd))
                                {
                                    colName = GetColumnNameFromColumnNumber(cInd);
                                }
                            }
                            else
                            {
                                colName = GetColumnNameByHeading(fileName, sheetName, str.TrimStart());
                            }
                            list.Add(colName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return list;
        }

        /// <summary>
        /// This method will return the columnname with columnheading provided
        /// </summary>
        /// <param name="colHeading"></param>
        /// <returns></returns>
        private string GetColumnNameByHeading(string fileName, string sheetName, string colHeading)
        {
            string colName = string.Empty;
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
                {
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                    IEnumerable<Sheet> Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

                    string relationshipId = Sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(relationshipId);
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    if (worksheetPart != null)
                    {
                        var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                        if (rows != null)
                        {
                            foreach (var row in rows)
                            {
                                var cols = row.Elements<Cell>();
                                foreach (var col in cols)
                                {
                                    string cellHead = GetCurrentCellValue(spreadSheet, col);
                                    if (cellHead == colHeading)
                                    {
                                        colName = RemoveIntegerFromColumnName(col.CellReference.InnerText);
                                        break;
                                    }
                                }
                                if (!string.IsNullOrEmpty(colName)) break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return colName;
        }

        /// <summary>
        /// This method will return the column name
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        private string GetColumnName(string fileName, string sheetName, string col)
        {
            string clName = string.Empty;
            try
            {
                if (col.Trim().StartsWith("#"))
                {
                    string cl = col.Trim().Replace("#", "");
                    int cIndex = 0;
                    if (int.TryParse(cl, out cIndex))
                    {
                        clName = GetColumnNameFromColumnNumber(cIndex);
                    }
                    else if (cl.Length == 1)
                    {
                        clName = cl;
                    }
                }
                else
                {
                    clName = GetColumnNameByHeading(fileName, sheetName, col.Trim());
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return clName;
        }

        /// <summary>
        /// This method is used to remove integer from columnName
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private string RemoveIntegerFromColumnName(string columnName)
        {
            string cName = string.Empty;
            try
            {
                columnName = columnName.Replace("#", "");
                foreach (char ch in columnName)
                {
                    int num = 0;
                    if (int.TryParse(Convert.ToString(ch), out num))
                    {
                        break;
                    }
                    else
                    {
                        cName = cName + Convert.ToString(ch);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return cName;
        }

        #endregion

        #region Insert Update Methods

        // Given a document name and text, 
        // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        public static bool InsertText(string fileName, string sheetName, uint row, string column, string value)
        {
            bool isInserted = false;
            try
            {
                // Open the document for editing.
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
                {
                    IEnumerable<Sheet> Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    string relationshipId = Sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(relationshipId);
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    bool isNewRow = false;
                    Row newRow = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == row);
                    if (newRow == null)
                    {
                        newRow = new Row();
                        newRow.RowIndex = row;
                        isNewRow = true;
                    }

                    var cell = newRow.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, column + row, true) == 0).FirstOrDefault();
                    if (cell == null)
                    {
                        cell = new Cell();
                        cell.CellReference = column + Convert.ToString(row);
                        cell.DataType = CellValues.String;
                    }
                    
                    // Set the value of cell A1.
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    if (isNewRow)
                    {
                        newRow.AppendChild(cell);
                        sheetData.AppendChild(newRow);
                    }
                    // Save the new worksheet.
                    worksheetPart.Worksheet.Save();
                    isInserted = true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isInserted;
        }
        
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        /// <summary>
        /// This method updates the cells for the row
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="updateValues"></param>
        private bool UpdateRowCells(string fileName, string sheetName, int rowIndex, string updateValues)
        {
            bool isUpdated = false;
            try
            {
                List<ExcelCellValues> listCols = GetColumnUpdateValues(fileName, sheetName, updateValues);
                if (listCols.Count > 0)
                {
                    // Open the document for editing.
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, true))
                    {
                        IEnumerable<Sheet> Sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                        string relationshipId = Sheets.First().Id.Value;
                        WorksheetPart worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(relationshipId);
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        if (worksheetPart != null)
                        {
                            var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
                            int curRowInd = 0;
                            foreach (var row in rows)
                            {
                                if (curRowInd == rowIndex)
                                {
                                    var cols = row.Elements<Cell>();
                                    foreach (var col in cols)
                                    {
                                        foreach (ExcelCellValues item in listCols)
                                        {
                                            if (IsColumnNameSame(item.CellName, col.CellReference.InnerText))
                                            {
                                                col.CellValue = new CellValue(item.CellUpdateValue);
                                                col.DataType = new EnumValue<CellValues>(CellValues.Number);
                                                isUpdated = true;
                                            }
                                        }
                                    }
                                }
                                curRowInd++;
                                if (isUpdated || curRowInd > rowIndex) break;
                            }

                            if (isUpdated)
                            {
                                // Save the worksheet.
                                worksheetPart.Worksheet.Save();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return isUpdated;
        }

        /// <summary>
        /// This method is used to Append the Row in Excel
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="firstName"></param>
        /// <param name="lastName"></param>
        /// <returns></returns>
        private Row AppendRowExcel(string fileName, string sheetName, List<ExcelCellValues> rowValueList)
        {
            Row contentRow = null;
            try
            {
                using (SpreadsheetDocument workbook = SpreadsheetDocument.Open(fileName, true))
                {
                    WorkbookPart workbookPart = workbook.WorkbookPart;
                    IEnumerable<Sheet> Sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
                    if (Sheets.Count() == 0)
                    {
                        // The specified worksheet does not exist.
                        return null;
                    }

                    string relationshipId = Sheets.First().Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)workbook.WorkbookPart.GetPartById(relationshipId);
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    int index = sheetData.Descendants<Row>().Count() + 1;
                    contentRow = GetRowToAppendInExcel(workbook, worksheetPart, index, rowValueList);
                    sheetData.AppendChild(contentRow);

                    workbookPart.Workbook.Save();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return contentRow;
        }

        /// <summary>
        /// This method is used to create the Row to append in excel
        /// </summary>
        /// <param name="rIndex"></param>
        /// <param name="firstName"></param>
        /// <param name="lastName"></param>
        /// <returns></returns>
        private static Row GetRowToAppendInExcel(SpreadsheetDocument document, WorksheetPart worksheetPart, int rIndex, List<ExcelCellValues> columnValueList)
        {
            Row r = new Row();
            try
            {
                r.RowIndex = (UInt32)rIndex;
                foreach (var colValue in columnValueList)
                {
                    if (!string.IsNullOrEmpty(colValue.CellUpdateValue))
                    {
                        Cell cell = new Cell();
                        cell.CellReference = colValue.CellName + Convert.ToString(rIndex);
                        cell.DataType = CellValues.String;
                        InlineString inlinefString = new InlineString();
                        Text txt = new Text();
                        txt.Text = colValue.CellUpdateValue;
                        inlinefString.AppendChild(txt);
                        cell.AppendChild(inlinefString);
                        r.AppendChild(cell);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return r;
        }
        
        #endregion

        #endregion
    }

    public class ExcelCellValues
    {
        public string CellName { get; set; }
        public string CellUpdateValue { get; set; }
    }
}
