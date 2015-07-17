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
        private Excel.Worksheet GetWorkSheetWithName(Excel.Workbook workBook, string WorksheetName, bool CreateNewIfNotFound = true)
        {
            foreach (var wksht in workBook.Worksheets)
            {
                Excel.Worksheet thisWorksheet = (Excel.Worksheet)wksht;
                if (thisWorksheet != null)
                {
                    if (thisWorksheet.Name == WorksheetName)
                    {
                        return thisWorksheet;
                    }
                }
            }
            if (CreateNewIfNotFound == true)
            {

                Excel.Worksheet newWorksheet = (Excel.Worksheet)workBook.Worksheets.Add();
                newWorksheet.Name = WorksheetName;
                return newWorksheet;
            }
            throw new Exception("Worksheet (tab) with the specified name not found in Excel file for input values. Please check input.");
        }

        public List<string> GetWorksheetsWithPrefix(string WorkbookPath, string prefix)
        {

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook xlwkbook = excelApplication.Workbooks.Open(WorkbookPath);
            List<string> wkshts = GetWorksheetsWithPrefix(xlwkbook, prefix);


            excelApplication.Quit();
            excelApplication.Dispose();
            return wkshts;
        }

        private List<string> GetWorksheetsWithPrefix(Excel.Workbook xlwkbook,  string prefix)
        {
            List<string> wkshts = new List<string>();

            foreach (var wksht in xlwkbook.Worksheets)
            {
                Excel.Worksheet thisWorksheet = (Excel.Worksheet)wksht;
                if (thisWorksheet != null)
                {
                    if (thisWorksheet.Name.StartsWith(prefix))
                    {
                        wkshts.Add(thisWorksheet.Name);
                    }
                }
            }

            return wkshts;
        }
    }
}
