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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wosad.ExcelClient
{
    public class ExcelInputReader : IInputReader
    {
        public ExcelInputReader(
            string WorkBookFileName, 
            string WorksheetName, 
            int startRow, 
            int StartColumn, 
            int NumberOfRows, 
            int NumberOfColumns)
        {
            this.WorkBookFileName= WorkBookFileName;
            this.WorksheetName=WorksheetName;
            this.startRow        =startRow;
            this.StartColumn     =StartColumn;
            this.NumberOfRows    =NumberOfRows;
            this.NumberOfColumns =NumberOfColumns;

        }

        List<ExcelEntry> retDict = new List<ExcelEntry>();
            string WorkBookFileName{get; set;}
            string WorksheetName   {get; set;}
            int startRow           {get; set;}
            int StartColumn        {get; set;}
            int NumberOfRows       {get; set;}
            int NumberOfColumns    {get; set;}


            bool excelDataRead;

            private List<ExcelEntry> lookupValues;

            private List<ExcelEntry> LookupValues
            {
                get
                {
                    return lookupValues;
                }

            }
            
            public string LookupInputValue(string ValueKey)
            {
                string Value = "0";
                if (excelDataRead == false)
                {
                    ReadWorksheet();
                }

                var foundVals = LookupValues.Where(x => x.ValueName == ValueKey).ToList();
                if (foundVals.Count != 0)
                {
                    Value = foundVals.First().ValueMagnitude;
                }

                return Value;
            }

            private void ReadWorksheet()
            {
                List<ExcelEntry> retDict = new List<ExcelEntry>();

                ExcelClientNetOffice ec = new ExcelClientNetOffice();
                List<string> varNames = ec.GetColumnOfValues(1, 1, WorkBookFileName, WorksheetName);
                List<string> varValues = ec.GetColumnOfValues(1, 2, WorkBookFileName, WorksheetName);
                if (varNames.Count == varValues.Count)
                {
                    for (int i = 0; i < varNames.Count; i++)
                    {
                        retDict.Add(new ExcelEntry() { ValueName = varNames[i], ValueMagnitude = varValues[i] });
                    }
                }
                else
                {
                    throw new Exception("Number of entries in the 1st column of Excel has to be the same as in second column.");
                }
                excelDataRead = true;
                lookupValues = retDict;
            }


    }
}
