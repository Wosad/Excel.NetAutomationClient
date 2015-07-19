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

using Caliburn.Micro;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Win32;
using Wosad.ExcelClient;

namespace Wosad.Excel.NetAutomationClient.Demo
{
    public class ExcelNetClientDemoViewModel : PropertyChangedBase
    {

        public ExcelNetClientDemoViewModel()
        {
            ExcelInputPath=@"C:\Temp\WorksheetInput.xlsx";
            ExcelOutputPath=@"C:\Temp\WorksheetOutput.xlsx";
            CellReadRow = 1;
            CellReadColumn = 1;
            CellWriteRow = 1;
            CellWriteColumn = 1;
            InputWorksheetName = "Sheet1";
            OutputWorksheetName = "Sheet1";
            CellTargetContent = "Hello";
        }


      #region Properties

        #region ExcelOutputPath 
		 
       private string _ExcelOutputPath;

        public string ExcelOutputPath
        {
            get { return _ExcelOutputPath; }
            set
            {
                _ExcelOutputPath = value;
                NotifyOfPropertyChange(() => ExcelOutputPath);
            }
        }
	    #endregion

        #region ExcelInputPath 
		 

            private string _ExcelInputPath;

            public string ExcelInputPath
            {
                get { return _ExcelInputPath; }
                set
                {
                    _ExcelInputPath = value;
                    NotifyOfPropertyChange(() => ExcelInputPath);
                }
            }
	    #endregion

        #region CellReadRow


            private int _CellReadRow;

            public int CellReadRow
            {
                get { return _CellReadRow; }
                set
                {
                    _CellReadRow = value;
                    NotifyOfPropertyChange(() => CellReadRow);
                }
            }
            #endregion

        #region CellReadColumn


            private int _CellReadColumn;

            public int CellReadColumn
            {
                get { return _CellReadColumn; }
                set
                {
                    _CellReadColumn = value;
                    NotifyOfPropertyChange(() => CellReadColumn);
                }
            }
            #endregion

        #region CellWriteRow


            private int _CellWriteRow;

            public int CellWriteRow
            {
                get { return _CellWriteRow; }
                set
                {
                    _CellWriteRow = value;
                    NotifyOfPropertyChange(() => CellWriteRow);
                }
            }
            #endregion

        #region CellWriteColumn


            private int _CellWriteColumn;

            public int CellWriteColumn
            {
                get { return _CellWriteColumn; }
                set
                {
                    _CellWriteColumn = value;
                    NotifyOfPropertyChange(() => CellWriteColumn);
                }
            }
            #endregion

        #region InputWorksheetName


            private string _InputWorksheetName;

            public string InputWorksheetName
            {
                get { return _InputWorksheetName; }
                set
                {
                    _InputWorksheetName = value;
                    NotifyOfPropertyChange(() => InputWorksheetName);
                }
            }
            #endregion

        #region OutputWorksheetName


            private string _OutputWorksheetName;

            public string OutputWorksheetName
            {
                get { return _OutputWorksheetName; }
                set
                {
                    _OutputWorksheetName = value;
                    NotifyOfPropertyChange(() => OutputWorksheetName);
                }
            }
            #endregion

        #region CellContent


            private string _CellContent;

            public string CellContent
            {
                get { return _CellContent; }
                set
                {
                    _CellContent = value;
                    NotifyOfPropertyChange(() => CellContent);
                }
            }
            #endregion

        #region CellTargetContent


            private string _CellTargetContent;

            public string CellTargetContent
            {
                get { return _CellTargetContent; }
                set
                {
                    _CellTargetContent = value;
                    NotifyOfPropertyChange(() => CellTargetContent);
                }
            }
            #endregion

        #endregion


      #region Commands

        public void GetExcelInputPath()
        {

            // Configure open file dialog box 
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = "InputFileName"; // Default file name 
            dlg.DefaultExt = ".xlsx"; // Default file extension 
            dlg.Filter = "Excel workbook (.xlsx)|*.xlsx"; // Filter files by extension 

            // Show open file dialog box 
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results 
            if (result == true)
            {
                ExcelInputPath = dlg.FileName;
            }
        }

        public void GetExcelOutputPath()
        {

            // Configure open file dialog box 
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = "OutputFileName"; // Default file name 
            dlg.DefaultExt = ".xlsx"; // Default file extension 
            dlg.Filter = "Excel workbook (.xlsx)|*.xlsx"; // Filter files by extension 

            // Show open file dialog box 
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results 
            if (result == true)
            {
                ExcelOutputPath = dlg.FileName;
            }
        } 


        public void ReadFile()
        {
            ExcelClientNetOffice ec = new ExcelClientNetOffice();
            var CellContentArr = ec.GetArrayOfValues(CellReadRow, CellReadColumn, 1, 1, ExcelInputPath, InputWorksheetName);
            CellContent = CellContentArr[0, 0];
        }

        public void WriteFile()
        {
            //ExcelClientNetOffice ec = new ExcelClientNetOffice();
            // ec.(CellReadRow, CellReadColumn, 1, 1, ExcelInputPath, InputWorksheetName);
            ExcelClientNetOffice ec = new ExcelClientNetOffice();
            string[,] Values = new string[1, 1];
            Values[0,0]= CellTargetContent;
            ExcelOutputEntry entry = new ExcelOutputEntry(){ 
                StartRow = CellWriteRow, 
                StartColumn = CellWriteColumn,
                Values = Values,
                Worksheet = OutputWorksheetName
            };
            List<ExcelOutputEntry> entries = new List<ExcelOutputEntry>(){entry};
            ec.DumpMultipleValueArrays(entries,ExcelOutputPath);
        }

        #endregion
    }
}
