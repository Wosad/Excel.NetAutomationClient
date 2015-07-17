//Copyright 2015 Wosad

//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at

//    http://www.apache.org/licenses/LICENSE-2.0

//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.

using System;
using System.Collections.Generic;
using Excel = NetOffice.ExcelApi;

namespace Wosad.ExcelClient
{
    public partial class ExcelClientNetOffice
    {

        public List<string> GetValuesFromListOfCells(List<ExcelAddress> addresses, string WorkbookPath, string WorksheetName)
        {
  

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, WorksheetName);


            string CellValue = null;
            List<string> Values = new List<string>();


            foreach (var addr in addresses)
            {
                Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[addr.Row, addr.Column];
                CellValue = Convert.ToString(CurrentRange.Cells.Value);
                Values.Add(CellValue);
            }
            return Values;
        }

        public List<List<string>> GetListOfValues (int StartRow, int StartColumn, int NumberOfColumns, string WorkbookPath, string WorksheetName)
        {
            int Row = StartRow;
            int Column = StartColumn;

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, WorksheetName);

            string CellValue = null;
            List<List<string>> Values = new List<List<string>>();

            while (CellValue != "")
            {
                Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[Row, StartColumn];
                CellValue = Convert.ToString(CurrentRange.Cells.Value);

                if (CellValue != "" && CellValue != null)
                {
                    List<string> currentRowValues = new List<string>();
                    for (int i = 0; i < NumberOfColumns; i++)
                    {
                        Excel.Range thisValueRange = (Excel.Range)xlsheet.Cells[Row, StartColumn+i];
                        string thisCellValue = Convert.ToString(thisValueRange.Cells.Value);
                        if (thisCellValue != null && thisCellValue !="")
                        {
                            currentRowValues.Add(thisCellValue);
                        }
                        else
                        {
                            currentRowValues.Add("");
                        }
                        
                    }
                    if (currentRowValues[0]!="" && currentRowValues[0]!=null)
                    {
                        Values.Add(currentRowValues);
                    }
                    
                }
                else
                {
                    CellValue = "";
                }
                Row++;
            }

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            return Values;
        }

        public List<string> GetColumnOfValues(int StartRow, int StartColumn, string WorkbookPath, string WorksheetName)
        {
            int Row = StartRow;
            int Column = StartColumn;

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, WorksheetName);

            string CellValue = null;
            List<string> Values = new List<string>();

            while (CellValue != "")
            {
                Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[Row, StartColumn];
                CellValue = Convert.ToString(CurrentRange.Cells.Value);
                Row++;
                if (CellValue != "" && CellValue != null)
                {

                    Values.Add(CellValue);
                }
                else
                {
                    CellValue = "";
                }

            }

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            return Values;
        }

        public string[,] GetArrayOfValues(
            int StartRow,
            int StartColumn,
            int numRows,
            int numColumns,
            string WorkbookPath,
            string WorksheetName)
        {
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, WorksheetName);


            string CellValue = null;
            List<string> Values = new List<string>();

            string[,] retArray = new string[numRows, numColumns];


            for (int i = 0; i < numRows; i++)
            {
                int arow = StartRow + i;

                for (int j = 0; j < numColumns; j++)
                {
                    int acol = StartColumn + j;

                    Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[arow, acol];
                    CellValue = Convert.ToString(CurrentRange.Cells.Value);
                    retArray[i, j] = CellValue;
                }
            }


            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            return retArray;
        }



        public Dictionary<string,string> GetValuesFromMultipleWorksheetsWithPrefix(string WorkbookPath, List<string> WorksheetPrefix, int Row, int Column)
        {
            Dictionary<string, string> values = new Dictionary<string, string>();
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            foreach (var pref in WorksheetPrefix)
            {
                List<string> wkshts = GetWorksheetsWithPrefix(xlwkbook, pref);

                foreach (var sheetName in wkshts)
                {
                    Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, sheetName);
                    Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[Row, Column];
                    string cVal = Convert.ToString(CurrentRange.Cells.Value);
                    values.Add(sheetName, cVal);
                } 
            }

            excelApplication.Quit();
            excelApplication.Dispose();
            return values;
        }
    }
}
