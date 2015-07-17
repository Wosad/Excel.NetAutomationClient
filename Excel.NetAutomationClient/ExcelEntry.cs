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


namespace Wosad.ExcelClient
{
    public class ExcelEntry
    {
        public ExcelEntry()
        {

        }
        public ExcelEntry(string ValueName, string ValueMagnitude)
        {
            this.ValueMagnitude = ValueMagnitude;
            this.ValueName = ValueName;
        }
        public string ValueName { get; set; }
        public string ValueMagnitude { get; set; }
    }
}
