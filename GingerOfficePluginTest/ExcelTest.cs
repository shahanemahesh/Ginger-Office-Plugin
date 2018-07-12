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

using GingerPlugInsNET.ActionsLib;
using GingerTestHelper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using StandAloneActions;
using System;
using System.Reflection;

namespace StandAloneActionsTest
{

    [TestClass]    
    public class ExcelTest
    {        

        static string EXCEL_FILE_NAME = TestResources.GetTestResourcesFile("test1.xlsx");


        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            TestResources.Assembly = Assembly.GetExecutingAssembly();
        }

        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            if (!System.IO.File.Exists(EXCEL_FILE_NAME))
            {
                Console.WriteLine("File not found: " + EXCEL_FILE_NAME);
            }
        }

        [TestMethod]
        public void ReadExcelCellRow3ColumnB()
        {            

            //Arrange
            ExcelAction x = new ExcelAction();
            GingerAction GA = new GingerAction("Excel");

            //Act
            x.ReadExcelCell(ref GA, EXCEL_FILE_NAME, "Sheet1", "#3", "#B");

            //Assert
            // Assert.AreEqual("Moshe", GA.Output.Values[0].ValueString, "Row 3 Col B = Moshe");
        }

        [TestMethod]
        public void ReadExcelCellRow3ColumnColumnCount()
        {
            //Arrange
            ExcelAction x = new ExcelAction();
            GingerAction GA = new GingerAction("Excel");

            //Act
            x.ReadExcelCell(ref GA, EXCEL_FILE_NAME, "Sheet1", "#3", "#B");

            //Assert
            // Assert.AreEqual(1, GA.Output.Values.Count);
        }

     
    }
}
