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
                    string col = "a";
                    string nRow = "3";
                    string txt1 = GetCellValue(FileName, "Sheet1", col + nRow);

                    GA.Output.Add("Value", txt1);
                    GA.ExInfo = "Read FileName: " + FileName + " row= " + row + " col= " + column;
                }
                else                                                        //row="First='Moshe'", column="ID"
                {
                    // List<string> colList = GetCoulmnsName(FileName, sheetName, column);
                    //ReadExcelRowWithCondition(GA, FileName, sheetName, row, colList);
                }
            }
            catch (Exception ex)
            {
                GA.AddError("ReadExcelCell", ex.StackTrace);
            }
        }

       

      
        
        #endregion

       
        
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

     
    }

    public class ExcelCellValues
    {
        public string CellName { get; set; }
        public string CellUpdateValue { get; set; }
    }
}