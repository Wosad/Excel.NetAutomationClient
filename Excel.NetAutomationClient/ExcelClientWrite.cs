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



using System.Collections.Generic;
using Excel = NetOffice.ExcelApi;

namespace Wosad.ExcelClient
{
    public partial class ExcelClientNetOffice
    {

        public void DumpMultipleValueArrays(List<ExcelOutputEntry> DataOutputEntries, string ExcelOutputPath)
        {

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(ExcelOutputPath);

            foreach (var entry in DataOutputEntries)
            {
                DumpArrayOfValues(entry.Values, entry.StartRow, entry.StartColumn, xlwkbook, entry.Worksheet);
            }

            xlwkbook.Save();
            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();
        }


        private void DumpArrayOfValues(string[,] Values, int StartRow, int StartColumn, Excel.Workbook xlwkbook, string WorksheetName)
        {

            Excel.Worksheet xlsheet = GetWorkSheetWithName(xlwkbook, WorksheetName);

            for (int i = 0; i < Values.GetLength(0); i++)
            {
                for (int j = 0; j < Values.GetLength(1); j++)
                {
                    int cRow = StartRow + i;
                    int cColumn = StartColumn + j;
                    Excel.Range CurrentRange = (Excel.Range)xlsheet.Cells[cRow, cColumn];
                    CurrentRange.Cells.Value = Values[i, j];
                }
            }


        }

    }
}
