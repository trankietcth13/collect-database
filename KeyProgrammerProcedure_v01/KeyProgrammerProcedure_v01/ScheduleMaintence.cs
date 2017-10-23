using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using System.IO;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OfficeOpenXml;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;

namespace GetDataOnWeb_v01
{
    public class ScheduleMaintence
    {

        public static ObjectYearMake[] arrayAllYearMake;

        public static ObjectModelTrim[] arrayAllModelTrim;

        public static ObjectEngineTrans[] arrayAllEngineTrans;

        public static ObjectModelYearId[] arrayAllModelYearId;

        public static ObjectSM[] arrayAllSM;

        public class ObjectYearMake
        {
            public string Make { get; set; }
            public ObjectYear[] Year { get; set; }
            public string NiceName { get; set; }
        }

        public class ObjectYear
        {
            public string Year { get; set; }
        }

        public class ObjectTrim
        {
            public string Trim { get; set; }
        }

        public class ObjectModelTrim
        {
            public string Model { get; set; }
            public ObjectTrim[] Trim { get; set; }
            public string ModelNiceName { get; set; }
        }

        public class ObjectEngineTrans
        {
            public string Engine { get; set; }

            public string Transmission { get; set; }
        }

        public class ObjectModelYearId
        {
            public string ModelYearId { get; set; }
            public string Trim { get; set; }
            public string Engine { get; set; }
            public string Body { get; set; }
        }

        public class ObjectSM
        {
            public string EngineCode { get; set; }
            public string IntervalMileage { get; set; }
            public string IntervalMonth { get; set; }
            public string Frequency { get; set; }
            public string Action { get; set; }
            public string Item { get; set; }
            public string ItemDes { get; set; }
            public string LaborUnits { get; set; }

            public string PartUnits { get; set; }
            public string DriveType { get; set; }

            public string PartCostPerUnit { get; set; }
        }

        public static string ActiononWeb(By byGetbtn, By byTrybtn, By byYear, By byState, By byCategory, By byMakeNiceName, By byModelNiceName, By byModelYearId, By byURLRequest, By byJson, bool yearcheck, bool statecheck, bool categorycheck, string txtyear, string state, string category, string MakeNicename, string ModelNiceName, string ModelYearId)
        {
            string jsonText = null;
            //Click button Get 
            IWebElement buttonGet = PropertiesCollection.driver.FindElement(byGetbtn);
            buttonGet.Click();

            System.Threading.Thread.Sleep(1000);

            //select all year
            if (yearcheck == true)
            {
                IWebElement elementinputYear = PropertiesCollection.driver.FindElement(byYear);
                elementinputYear.Clear();
                elementinputYear.SendKeys(txtyear);
                System.Threading.Thread.Sleep(1000);
            }

            //select state
            if (statecheck == true)
            {
                IWebElement elemenState = PropertiesCollection.driver.FindElement(byState);
                elemenState.AsDropDown().SelectByValue(state);
                System.Threading.Thread.Sleep(1000);
            }

            //select Category
            if (categorycheck == true)
            {
                IWebElement elemenCategory = PropertiesCollection.driver.FindElement(byCategory);
                elemenCategory.AsDropDown().SelectByValue(category);
                System.Threading.Thread.Sleep(1000);
            }

            //enter MakeNicename
            if (!(String.IsNullOrEmpty(MakeNicename)))
            {
                IWebElement elemenMakeNiceName = PropertiesCollection.driver.FindElement(byMakeNiceName);
                elemenMakeNiceName.Clear();
                elemenMakeNiceName.SendKeys(MakeNicename);
                System.Threading.Thread.Sleep(1000);
            }

            //enter ModelNicename
            if (!(String.IsNullOrEmpty(ModelNiceName)))
            {
                IWebElement elemenModelNiceName = PropertiesCollection.driver.FindElement(byModelNiceName);
                elemenModelNiceName.Clear();
                elemenModelNiceName.SendKeys(ModelNiceName);
                System.Threading.Thread.Sleep(1000);
            }

            //enter ModelYearId
            if (!(String.IsNullOrEmpty(ModelYearId)))
            {
                IWebElement elemenModelYearId = PropertiesCollection.driver.FindElement(byModelYearId);
                elemenModelYearId.Clear();
                elemenModelYearId.SendKeys(ModelYearId);
                System.Threading.Thread.Sleep(1000);

            }
            //Click button Try it to get source
            IWebElement btnTryit = PropertiesCollection.driver.FindElement(byTrybtn);
            btnTryit.Click();

            System.Threading.Thread.Sleep(4000);

            //get json URL
            IWebElement elementRequestURL = PropertiesCollection.driver.FindElement(byURLRequest);

            System.Threading.Thread.Sleep(2000);
            // open a new tab and set the context
            ChromeDriver drivertmp = (ChromeDriver)PropertiesCollection.driver;
            drivertmp.ExecuteScript(String.Format("window.open('{0}', 'tab2');", elementRequestURL.Text));
            drivertmp.SwitchTo().Window("tab2");

            IWebElement body = PropertiesCollection.driver.FindElement(By.TagName("body"));
            System.Threading.Thread.Sleep(2000);
            //body.SendKeys(Keys.Control + "t");

            //CommonMethods.SwitchToWindow(driver => driver.Title == "Title of your new tab");

            //PropertiesCollection.driver.Navigate().GoToUrl(elementRequestURL.Text);

            System.Threading.Thread.Sleep(3000);

            if (CommonMethods.IsElementPresent(byJson))
            {
                //get all year and make 
                IWebElement elementjson = PropertiesCollection.driver.FindElement(byJson);
                CommonMethods.DoWait(1000);
                jsonText = elementjson.Text;
                jsonText = "[" + jsonText + "]";
               
            }
            drivertmp.ExecuteScript(String.Format("window.close();"));
           
            //jsonText = jsonText.Remove(jsonText.Length - numCharEndtoDel, numCharEndtoDel);
            //jsonText = jsonText.Remove(0, numCharBegintoDel);

            //back previous page
            //PropertiesCollection.driver.Navigate().Back();

            return jsonText;
        }

        public static void HandleonWeb()
        {
            string APIKey = "sv68rfxgdc7qxvc9payea3fq";

            //row of data
            int rowYMME = 2;
            int rowDB = 2;

            string path = Environment.CurrentDirectory;
            string foldername = Path.Combine(path, "outputScheduleMaintenance.xlsx");

            //MyBook.SaveAs(foldername);
            System.Threading.Thread.Sleep(500);

            //get all Makes and all years of each make
            //GetAllYearMake(APIKey);
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                workbook.Worksheets.Add("YMME");
                var MySheet = workbook.Worksheets[1];
                workbook.Worksheets.Add("Database");
                var MySheet2 = workbook.Worksheets[2];

                //create header of YMME sheet
                MySheet.Cells[1, 1].Value = "YEAR";
                MySheet.Cells[1, 2].Value = "MAKE";
                MySheet.Cells[1, 3].Value = "MODEL";
                MySheet.Cells[1, 4].Value = "ENGINE";
                MySheet.Cells[1, 5].Value = "TRANSMISSION";

                //create header of Database sheet
                MySheet2.Cells[1, 1].Value = "YEAR";
                MySheet2.Cells[1, 2].Value = "MAKE";
                MySheet2.Cells[1, 3].Value = "MODEL";
                MySheet2.Cells[1, 4].Value = "TRIM";
                MySheet2.Cells[1, 5].Value = "ENGINE";
                MySheet2.Cells[1, 6].Value = "BODY";
                MySheet2.Cells[1, 7].Value = "engineCode";
                MySheet2.Cells[1, 8].Value = "intervalMileage";
                MySheet2.Cells[1, 9].Value = "intervalMonth";
                MySheet2.Cells[1, 10].Value = "frequency";
                MySheet2.Cells[1, 11].Value = "action";
                MySheet2.Cells[1, 12].Value = "item";
                MySheet2.Cells[1, 13].Value = "itemDescription";
                MySheet2.Cells[1, 14].Value = "laborUnits";
                MySheet2.Cells[1, 15].Value = "partUnits";
                MySheet2.Cells[1, 16].Value = "driveType";
                MySheet2.Cells[1, 17].Value = "partCostPerUnit ";

                package.SaveAs(new System.IO.FileInfo(foldername));

                //by array Make
                for (int i = 0; i < arrayAllYearMake.Count(); i++)
                {

                    //by array Year
                    for (int j = 0; j < arrayAllYearMake[i].Year.Count(); j++)
                    {

                        //get all Models and all Trims of each make via MakeNicename and Year
                        GetModelandTrimbyMakeNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllYearMake[i].Year[j].Year);

                        //by array Model
                        for (int k = 0; k < arrayAllModelTrim.Count(); k++)
                        {

                            //get all Engines and Transmissions via MakeNicename, ModelNicename and Year
                            GetEngineandTransbyMakeandModelNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllModelTrim[k].ModelNiceName, arrayAllYearMake[i].Year[j].Year);

                            //by array Engine
                            for (int indexEngine = 0; indexEngine < arrayAllEngineTrans.Count(); indexEngine++)
                            {
                                if (String.IsNullOrEmpty(arrayAllEngineTrans[indexEngine].Engine))
                                {
                                    continue;
                                }
                                if (indexEngine != 0 && (String.IsNullOrEmpty(arrayAllEngineTrans[indexEngine - 1].Engine) || arrayAllEngineTrans[indexEngine].Engine.Equals(arrayAllEngineTrans[indexEngine - 1].Engine)))
                                {
                                    continue;
                                }
                                //write year
                                MySheet.Cells[rowYMME, 1].Value = arrayAllYearMake[i].Year[j].Year;
                                //write make            
                                MySheet.Cells[rowYMME, 2].Value = arrayAllYearMake[i].Make;
                                //write model            
                                MySheet.Cells[rowYMME, 3].Value = arrayAllModelTrim[k].Model;
                                //write trim             
                                MySheet.Cells[rowYMME, 4].Value = arrayAllEngineTrans[indexEngine].Engine;
                                //write engine          
                                MySheet.Cells[rowYMME, 5].Value = arrayAllEngineTrans[indexEngine].Transmission;

                                rowYMME++;
                                package.Save();
                            }

                            //get all ModelYearIds via  MakeNicename and ModelNicename
                            GetModelYearIdbyMakeandModelNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllModelTrim[k].ModelNiceName);

                            //by array ModelYearId
                            for (int m = 0; m < arrayAllModelYearId.Count(); m++)
                            {
                                //get all ScheduleMaintenance via ModelYearId
                                GetScheduleMantenancebyModelYearId(APIKey, arrayAllModelYearId[m].ModelYearId);

                                //by array ScheduleMaintenance
                                for (int sm = 0; sm < arrayAllSM.Count(); sm++)
                                {
                                    //write year
                                    MySheet2.Cells[rowDB, 1].Value = arrayAllYearMake[i].Year[j].Year;
                                    //write make            
                                    MySheet2.Cells[rowDB, 2].Value = arrayAllYearMake[i].Make;
                                    //write model          
                                    MySheet2.Cells[rowDB, 3].Value = arrayAllModelTrim[k].Model;
                                    //write trim 
                                    MySheet2.Cells[rowDB, 4].Value = arrayAllModelYearId[m].Trim;
                                    //write engine
                                    MySheet2.Cells[rowDB, 5].Value = arrayAllModelYearId[m].Engine;
                                    //write body            
                                    MySheet2.Cells[rowDB, 6].Value = arrayAllModelYearId[m].Body;
                                    //write engineCode      
                                    MySheet2.Cells[rowDB, 7].Value = arrayAllSM[sm].EngineCode;
                                    //write intervalMileage 
                                    MySheet2.Cells[rowDB, 8].Value = arrayAllSM[sm].IntervalMileage;
                                    //write intervalMonth   
                                    MySheet2.Cells[rowDB, 9].Value = arrayAllSM[sm].IntervalMonth;
                                    //write frequency       
                                    MySheet2.Cells[rowDB, 10].Value = arrayAllSM[sm].Frequency;
                                    //write action          
                                    MySheet2.Cells[rowDB, 11].Value = arrayAllSM[sm].Action;
                                    //write item            
                                    MySheet2.Cells[rowDB, 12].Value = arrayAllSM[sm].Item;
                                    //write itemDescription 
                                    MySheet2.Cells[rowDB, 13].Value = arrayAllSM[sm].ItemDes;
                                    //write laborUnits      
                                    MySheet2.Cells[rowDB, 14].Value = arrayAllSM[sm].LaborUnits;
                                    //write partUnits       
                                    MySheet2.Cells[rowDB, 15].Value = arrayAllSM[sm].PartUnits;
                                    //write driveType       
                                    MySheet2.Cells[rowDB, 16].Value = arrayAllSM[sm].DriveType;
                                    //write partCostPerUnit 
                                    MySheet2.Cells[rowDB, 17].Value = arrayAllSM[sm].PartCostPerUnit;

                                    rowDB++;
                                    package.Save();
                                }
                            }
                        }
                    }
                }
                package.Save();

            }

            //MyBook.Save();
            //MyBook.Close();
            //MyApp.Quit();
        }

        public static void HandleonWebforSelectedMakeYear(string make, string year)
        {
            string APIKey = "sv68rfxgdc7qxvc9payea3fq";

            //row of data
            int rowYMME = 2;
            int rowDB = 2;
            int rowIdFail = 2;

            string path = Environment.CurrentDirectory;
            string namefile = "outputScheduleMaintenance_" + make + ".xlsx";
            string foldername = Path.Combine(path, namefile);

            //MyBook.SaveAs(foldername);
            System.Threading.Thread.Sleep(500);

            //get all Makes and all years of each make
            //GetAllYearMake(APIKey);

            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                workbook.Worksheets.Add("YMME");
                var MySheet = workbook.Worksheets[1];
                workbook.Worksheets.Add("Database");
                var MySheet2 = workbook.Worksheets[2];
                workbook.Worksheets.Add("Id Fail");
                var MySheet3 = workbook.Worksheets[3];

                //create header of YMME sheet
                MySheet.Cells[1, 1].Value = "YEAR";
                MySheet.Cells[1, 2].Value = "MAKE";
                MySheet.Cells[1, 3].Value = "MODEL";
                MySheet.Cells[1, 4].Value = "ENGINE";
                MySheet.Cells[1, 5].Value = "TRANSMISSION";

                //create header of Database sheet
                MySheet2.Cells[1, 1].Value = "YEAR";
                MySheet2.Cells[1, 2].Value = "MAKE";
                MySheet2.Cells[1, 3].Value = "MODEL";
                MySheet2.Cells[1, 4].Value = "TRIM";
                MySheet2.Cells[1, 5].Value = "ENGINE";
                MySheet2.Cells[1, 6].Value = "BODY";
                MySheet2.Cells[1, 7].Value = "engineCode";
                MySheet2.Cells[1, 8].Value = "intervalMileage";
                MySheet2.Cells[1, 9].Value = "intervalMonth";
                MySheet2.Cells[1, 10].Value = "frequency";
                MySheet2.Cells[1, 11].Value = "action";
                MySheet2.Cells[1, 12].Value = "item";
                MySheet2.Cells[1, 13].Value = "itemDescription";
                MySheet2.Cells[1, 14].Value = "laborUnits";
                MySheet2.Cells[1, 15].Value = "partUnits";
                MySheet2.Cells[1, 16].Value = "driveType";
                MySheet2.Cells[1, 17].Value = "partCostPerUnit ";

                // //create header of ID Fail sheet
                MySheet3.Cells[1, 1].Value = "Id Name";
                MySheet3.Cells[1, 2].Value = "Id Value";

                package.SaveAs(new System.IO.FileInfo(foldername));

                //by array Make
                for (int i = 0; i < arrayAllYearMake.Count(); i++)
                {
                    if (String.IsNullOrEmpty(year))
                    {
                        if (arrayAllYearMake[i].Make.Contains(make))
                        {
                            //by array Year
                            for (int j = 0; j < arrayAllYearMake[i].Year.Count(); j++)
                            {

                                //get all Models and all Trims of each make via MakeNicename and Year
                                GetModelandTrimbyMakeNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllYearMake[i].Year[j].Year);

                                //by array Model
                                for (int k = 0; k < arrayAllModelTrim.Count(); k++)
                                {

                                    //get all Engines and Transmissions via MakeNicename, ModelNicename and Year
                                    bool Enginecheck = GetEngineandTransbyMakeandModelNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllModelTrim[k].ModelNiceName, arrayAllYearMake[i].Year[j].Year);
                                    if (Enginecheck == false)
                                    {
                                        MySheet3.Cells[rowIdFail, 1].Value = "ModelNiceName";
                                        MySheet3.Cells[rowIdFail, 2].Value = arrayAllModelTrim[k].ModelNiceName;
                                        MySheet3.Cells[rowIdFail, 3].Value = arrayAllYearMake[i].NiceName;
                                        rowIdFail++;

                                        package.Save();
                                        continue;
                                    }
                                    //by array Engine
                                    for (int indexEngine = 0; indexEngine < arrayAllEngineTrans.Count(); indexEngine++)
                                    {
                                        if (String.IsNullOrEmpty(arrayAllEngineTrans[indexEngine].Engine))
                                        {
                                            continue;
                                        }
                                        if (indexEngine != 0 && (String.IsNullOrEmpty(arrayAllEngineTrans[indexEngine - 1].Engine) || arrayAllEngineTrans[indexEngine].Engine.Equals(arrayAllEngineTrans[indexEngine - 1].Engine)))
                                        {
                                            continue;
                                        }
                                        //write year
                                        MySheet.Cells[rowYMME, 1].Value = arrayAllYearMake[i].Year[j].Year;
                                        //write make  
                                        MySheet.Cells[rowYMME, 2].Value = arrayAllYearMake[i].Make;
                                        //write model 
                                        MySheet.Cells[rowYMME, 3].Value = arrayAllModelTrim[k].Model;
                                        //write trim  
                                        MySheet.Cells[rowYMME, 4].Value = arrayAllEngineTrans[indexEngine].Engine;
                                        //write engine
                                        MySheet.Cells[rowYMME, 5].Value = arrayAllEngineTrans[indexEngine].Transmission;

                                        rowYMME++;
                                        package.Save();
                                    }

                                    //get all ModelYearIds via  MakeNicename and ModelNicename
                                    GetModelYearIdbyMakeandModelNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllModelTrim[k].ModelNiceName);

                                    //by array ModelYearId
                                    for (int m = 0; m < arrayAllModelYearId.Count(); m++)
                                    {
                                        //get all ScheduleMaintenance via ModelYearId
                                        bool SM = GetScheduleMantenancebyModelYearId(APIKey, arrayAllModelYearId[m].ModelYearId);

                                        if (SM == false)
                                        {
                                            MySheet3.Cells[rowIdFail, 1].Value = "ModelYearId";
                                            MySheet3.Cells[rowIdFail, 2].Value = arrayAllModelYearId[m].ModelYearId;
                                            rowIdFail++;

                                            package.Save();
                                            continue;
                                        }

                                        //by array ScheduleMaintenance
                                        for (int sm = 0; sm < arrayAllSM.Count(); sm++)
                                        {
                                            //write year
                                            MySheet2.Cells[rowDB, 1].Value = arrayAllYearMake[i].Year[j].Year;
                                            //write make
                                            MySheet2.Cells[rowDB, 2].Value = arrayAllYearMake[i].Make;
                                            //write model
                                            MySheet2.Cells[rowDB, 3].Value = arrayAllModelTrim[k].Model;
                                            //write trim
                                            MySheet2.Cells[rowDB, 4].Value = arrayAllModelYearId[m].Trim;
                                            //write engine
                                            MySheet2.Cells[rowDB, 5].Value = arrayAllModelYearId[m].Engine;
                                            //write body
                                            MySheet2.Cells[rowDB, 6].Value = arrayAllModelYearId[m].Body;
                                            //write engineCode
                                            MySheet2.Cells[rowDB, 7].Value = arrayAllSM[sm].EngineCode;
                                            //write intervalMileage             
                                            MySheet2.Cells[rowDB, 8].Value = arrayAllSM[sm].IntervalMileage;
                                            //write intervalMonth            
                                            MySheet2.Cells[rowDB, 9].Value = arrayAllSM[sm].IntervalMonth;
                                            //write frequency
                                            MySheet2.Cells[rowDB, 10].Value = arrayAllSM[sm].Frequency;
                                            //write action            
                                            MySheet2.Cells[rowDB, 11].Value = arrayAllSM[sm].Action;
                                            //write item
                                            MySheet2.Cells[rowDB, 12].Value = arrayAllSM[sm].Item;
                                            //write itemDescription             
                                            MySheet2.Cells[rowDB, 13].Value = arrayAllSM[sm].ItemDes;
                                            //write laborUnits
                                            MySheet2.Cells[rowDB, 14].Value = arrayAllSM[sm].LaborUnits;
                                            //write partUnits             
                                            MySheet2.Cells[rowDB, 15].Value = arrayAllSM[sm].PartUnits;
                                            //write driveType
                                            MySheet2.Cells[rowDB, 16].Value = arrayAllSM[sm].DriveType;
                                            //write partCostPerUnit
                                            MySheet2.Cells[rowDB, 17].Value = arrayAllSM[sm].PartCostPerUnit;

                                            rowDB++;
                                            package.Save();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (arrayAllYearMake[i].Make.Contains(make))
                        {
                            //by array Year
                            for (int j = 0; j < arrayAllYearMake[i].Year.Count(); j++)
                            {
                                if (arrayAllYearMake[i].Year[j].Year.Contains(year))
                                {
                                    //get all Models and all Trims of each make via MakeNicename and Year
                                    GetModelandTrimbyMakeNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllYearMake[i].Year[j].Year);

                                    //by array Model
                                    for (int k = 0; k < arrayAllModelTrim.Count(); k++)
                                    {

                                        //get all Engines and Transmissions via MakeNicename, ModelNicename and Year
                                        bool Enginecheck = GetEngineandTransbyMakeandModelNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllModelTrim[k].ModelNiceName, arrayAllYearMake[i].Year[j].Year);
                                        if (Enginecheck == false)
                                        {
                                            MySheet3.Cells[rowIdFail, 1].Value = "ModelNiceName";
                                            MySheet3.Cells[rowIdFail, 2].Value = arrayAllModelTrim[k].ModelNiceName;
                                            MySheet3.Cells[rowIdFail, 3].Value = arrayAllYearMake[i].NiceName;
                                            rowIdFail++;

                                            package.Save();
                                            continue;
                                        }
                                        //by array Engine
                                        for (int indexEngine = 0; indexEngine < arrayAllEngineTrans.Count(); indexEngine++)
                                        {
                                            if (String.IsNullOrEmpty(arrayAllEngineTrans[indexEngine].Engine))
                                            {
                                                continue;
                                            }
                                            if (indexEngine != 0 && (String.IsNullOrEmpty(arrayAllEngineTrans[indexEngine - 1].Engine) || arrayAllEngineTrans[indexEngine].Engine.Equals(arrayAllEngineTrans[indexEngine - 1].Engine)))
                                            {
                                                continue;
                                            }
                                            //write year
                                            MySheet.Cells[rowYMME, 1].Value = arrayAllYearMake[i].Year[j].Year;
                                            //write make  
                                            MySheet.Cells[rowYMME, 2].Value = arrayAllYearMake[i].Make;
                                            //write model 
                                            MySheet.Cells[rowYMME, 3].Value = arrayAllModelTrim[k].Model;
                                            //write trim  
                                            MySheet.Cells[rowYMME, 4].Value = arrayAllEngineTrans[indexEngine].Engine;
                                            //write engine
                                            MySheet.Cells[rowYMME, 5].Value = arrayAllEngineTrans[indexEngine].Transmission;

                                            rowYMME++;
                                            package.Save();
                                        }

                                        //get all ModelYearIds via  MakeNicename and ModelNicename
                                        GetModelYearIdbyMakeandModelNicename(APIKey, arrayAllYearMake[i].NiceName, arrayAllModelTrim[k].ModelNiceName);

                                        //by array ModelYearId
                                        for (int m = 0; m < arrayAllModelYearId.Count(); m++)
                                        {
                                            //get all ScheduleMaintenance via ModelYearId
                                            bool SM = GetScheduleMantenancebyModelYearId(APIKey, arrayAllModelYearId[m].ModelYearId);
                                            if (SM == false)
                                            {
                                                MySheet3.Cells[rowIdFail, 1].Value = "ModelYearId";
                                                MySheet3.Cells[rowIdFail, 2].Value = arrayAllModelYearId[m].ModelYearId;
                                                rowIdFail++;

                                                package.Save();
                                                continue;
                                            }
                                            //by array ScheduleMaintenance
                                            for (int sm = 0; sm < arrayAllSM.Count(); sm++)
                                            {
                                                //write year
                                                MySheet2.Cells[rowDB, 1].Value = arrayAllYearMake[i].Year[j].Year;
                                                //write make
                                                MySheet2.Cells[rowDB, 2].Value = arrayAllYearMake[i].Make;
                                                //write model
                                                MySheet2.Cells[rowDB, 3].Value = arrayAllModelTrim[k].Model;
                                                //write trim
                                                MySheet2.Cells[rowDB, 4].Value = arrayAllModelYearId[m].Trim;
                                                //write engine
                                                MySheet2.Cells[rowDB, 5].Value = arrayAllModelYearId[m].Engine;
                                                //write body
                                                MySheet2.Cells[rowDB, 6].Value = arrayAllModelYearId[m].Body;
                                                //write engineCode
                                                MySheet2.Cells[rowDB, 7].Value = arrayAllSM[sm].EngineCode;
                                                //write intervalMileage                    
                                                MySheet2.Cells[rowDB, 8].Value = arrayAllSM[sm].IntervalMileage;
                                                //write intervalMonth                   
                                                MySheet2.Cells[rowDB, 9].Value = arrayAllSM[sm].IntervalMonth;
                                                //write frequency
                                                MySheet2.Cells[rowDB, 10].Value = arrayAllSM[sm].Frequency;
                                                //write action            
                                                MySheet2.Cells[rowDB, 11].Value = arrayAllSM[sm].Action;
                                                //write item
                                                MySheet2.Cells[rowDB, 12].Value = arrayAllSM[sm].Item;
                                                //write itemDescription             
                                                MySheet2.Cells[rowDB, 13].Value = arrayAllSM[sm].ItemDes;
                                                //write laborUnits
                                                MySheet2.Cells[rowDB, 14].Value = arrayAllSM[sm].LaborUnits;
                                                //write partUnits             
                                                MySheet2.Cells[rowDB, 15].Value = arrayAllSM[sm].PartUnits;
                                                //write driveType
                                                MySheet2.Cells[rowDB, 16].Value = arrayAllSM[sm].DriveType;
                                                //write partCostPerUnit
                                                MySheet2.Cells[rowDB, 17].Value = arrayAllSM[sm].PartCostPerUnit;

                                                rowDB++;
                                                package.Save();
                                            }
                                        }
                                    }
                                }

                            }
                        }
                    }

                }
                package.Save();
            }

            //MyBook.Save();
            //MyBook.Close();
            //MyApp.Quit();
        }

        public static void GetAllYearMake(string APIKey)
        {
            //navigate to URL
            PropertiesCollection.NavigatetoURL("http://edmunds.mashery.com/io-docs");
            //enter API key
            IWebElement elementAPIKey = PropertiesCollection.driver.FindElement(By.Id("apiKey"));
            CommonMethods.DoWait(1000);
            //System.Threading.Thread.Sleep(1000);
            elementAPIKey.SendKeys(APIKey);

            //Click button Get 
            string btnGetAllYearMake = "html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/div/span[1]";
            string elementinputYear = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/table/tbody/tr[2]/td[2]/input";
            string elementState = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/table/tbody/tr[1]/td[2]/select";
            string btnTryitAllYearsMakes = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/input[5]";
            string elementRequestURL = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/div[1]/pre[1]";
            string elementjson = "/html/body/pre";
            string jsonText = ActiononWeb(By.XPath(btnGetAllYearMake), By.XPath(btnTryitAllYearsMakes), By.XPath(elementinputYear), By.XPath(elementState), null, By.XPath(""), By.XPath(""), By.XPath(""), By.XPath(elementRequestURL), By.XPath(elementjson), true, true, false, "", "used", "", "", "", "");

            //objects is all make
            var objects = JArray.Parse(jsonText);
            int indexofMake = 0;


            foreach (JObject root in objects)
            {
                foreach (KeyValuePair<String, JToken> properties in root)
                {
                    if (properties.Key.Contains("Count"))
                    {
                        break;
                    }
                    arrayAllYearMake = new ObjectYearMake[properties.Value.Count()];
                    foreach (JObject memPro in properties.Value)
                    {
                        arrayAllYearMake[indexofMake] = new ObjectYearMake();
                        //root is make with id, name, nicename and models[]
                        foreach (KeyValuePair<String, JToken> app in memPro)
                        {
                            switch (app.Key.ToString())
                            {
                                case "name":
                                    arrayAllYearMake[indexofMake].Make = app.Value.ToString();
                                    break;
                                case "niceName":
                                    arrayAllYearMake[indexofMake].NiceName = app.Value.ToString();
                                    break;
                                case "models":
                                    foreach (JObject model in app.Value)
                                    {
                                        foreach (KeyValuePair<String, JToken> memModel in model)
                                        {
                                            if (memModel.Key.Contains("years"))
                                            {
                                                int indexofYear = 0;
                                                arrayAllYearMake[indexofMake].Year = new ObjectYear[memModel.Value.Count()];
                                                foreach (JObject year in memModel.Value)
                                                {
                                                    foreach (KeyValuePair<String, JToken> memyear in year)
                                                    {
                                                        arrayAllYearMake[indexofMake].Year[indexofYear] = new ObjectYear();
                                                        if (memyear.Key.Contains("year"))
                                                        {
                                                            arrayAllYearMake[indexofMake].Year[indexofYear].Year = memyear.Value.ToString();
                                                            indexofYear++;
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }
                                    indexofMake++;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        public static void GetModelandTrimbyMakeNicename(string APIKey, string Makenicename, string year)
        {
            //navigate to URL
            //PropertiesCollection.NavigatetoURL("http://edmunds.mashery.com/io-docs");
            //enter API key
            //IWebElement elementAPIKey = PropertiesCollection.driver.FindElement(By.Id("apiKey"));
            CommonMethods.DoWait(1000);
            //elementAPIKey.SendKeys(APIKey);
            
            //Click button Get 
            string btnGetAllModels = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/div/span[1]";
            string elementinputYear = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/form/table/tbody/tr[3]/td[2]/input";
            string elementState = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/form/table/tbody/tr[2]/td[2]/select";
            string elementCategory = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/form/table/tbody/tr[5]/td[2]/select";
            string elementMakeNicename = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/form/table/tbody/tr[1]/td[2]/input";
            string btnTryIt = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/form/input[5]";
            string elementRequestURL = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[2]/ul/li[2]/form/div[1]/pre[1]";
            string elementjson = "/html/body/pre";
            string jsonText = ActiononWeb(By.XPath(btnGetAllModels), By.XPath(btnTryIt), By.XPath(elementinputYear), By.XPath(elementState), By.XPath(elementCategory), By.XPath(elementMakeNicename), By.XPath(""), By.XPath(""), By.XPath(elementRequestURL), By.XPath(elementjson), true, true, true, year, "used", "", Makenicename, "", "");

            if(jsonText != null)
            {
                //objects is all make
                var objects = JArray.Parse(jsonText);
                int indexofModel = 0;

                foreach (JObject root in objects)
                {
                    foreach (KeyValuePair<String, JToken> properties in root)
                    {
                        if (properties.Key.Contains("Count"))
                        {
                            break;
                        }
                        arrayAllModelTrim = new ObjectModelTrim[properties.Value.Count()];
                        foreach (JObject memPro in properties.Value)
                        {
                            if (indexofModel >= arrayAllModelTrim.Count())
                            {
                                break;
                            }
                            arrayAllModelTrim[indexofModel] = new ObjectModelTrim();
                            //root is make with id, name, nicename and models[]
                            foreach (KeyValuePair<String, JToken> app in memPro)
                            {
                                //app is member of root with 2 value: Key and value of Key
                                if (app.Key.Contains("name"))
                                {
                                    arrayAllModelTrim[indexofModel].Model = app.Value.ToString();

                                }

                                if (app.Key.Contains("niceName"))
                                {
                                    arrayAllModelTrim[indexofModel].ModelNiceName = app.Value.ToString();
                                }

                                if (app.Key.Contains("years"))
                                {
                                    if (indexofModel >= arrayAllModelTrim.Count())
                                    {
                                        break;
                                    }

                                    //foreach (JProperty model in app.Value)
                                    //{
                                    //    if (model.Name.Contains("styles"))
                                    //    {
                                    //        int indexofTrim = 0;
                                    //        arrayAllModelTrim[indexofModel].Trim = new ObjectTrim[model.Value.Count()];
                                    //        foreach(JProperty trim in model.Value)
                                    //        {
                                    //            if (indexofTrim >= arrayAllModelTrim[indexofModel].Trim.Count())
                                    //            {
                                    //                break;
                                    //            }

                                    //            arrayAllModelTrim[indexofModel].Trim[indexofTrim] = new ObjectTrim();
                                    //            if (trim.Name.Contains("name"))
                                    //            {
                                    //                arrayAllModelTrim[indexofModel].Trim[indexofTrim].Trim = trim.Value.ToString();
                                    //                indexofTrim++;
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    foreach (JObject model in app.Value)
                                    {
                                        foreach (KeyValuePair<String, JToken> memModel in model)
                                        {
                                            if (memModel.Key.Contains("styles"))
                                            {
                                                int indexofTrim = 0;
                                                arrayAllModelTrim[indexofModel].Trim = new ObjectTrim[memModel.Value.Count()];
                                                foreach (JObject trim in memModel.Value)
                                                {
                                                    foreach (KeyValuePair<String, JToken> memtrim in trim)
                                                    {
                                                        if (indexofTrim >= arrayAllModelTrim[indexofModel].Trim.Count())
                                                        {
                                                            break;
                                                        }

                                                        arrayAllModelTrim[indexofModel].Trim[indexofTrim] = new ObjectTrim();
                                                        if (memtrim.Key.Contains("name"))
                                                        {
                                                            arrayAllModelTrim[indexofModel].Trim[indexofTrim].Trim = memtrim.Value.ToString();
                                                            indexofTrim++;
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }
                                    indexofModel++;
                                }
                            }
                        }
                    }
                }
            }
            
        }

        public static bool GetEngineandTransbyMakeandModelNicename(string APIKey, string Makenicename, string Modelnicename, string year)
        {
            //navigate to URL
            PropertiesCollection.NavigatetoURL("http://edmunds.mashery.com/io-docs");
            if (CommonMethods.IsElementPresent(By.Id("apiKey")))
            {
                //enter API key
                IWebElement elementAPIKey = PropertiesCollection.driver.FindElement(By.Id("apiKey"));
                CommonMethods.DoWait(1000);
                elementAPIKey.SendKeys(APIKey);
                //Click button Get 

                string btnGetAllEngineandTrans = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/div/span[1]";
                string elementinputYear = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/table/tbody/tr[3]/td[2]/input";
                string elementState = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/table/tbody/tr[4]/td[2]/select";
                string elementCategory = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/table/tbody/tr[6]/td[2]/select";
                string elementMakeNicename = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/table/tbody/tr[1]/td[2]/input";
                string elementModelNicename = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/table/tbody/tr[2]/td[2]/input";
                string btnTryIt = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/input[5]";
                string elementRequestURL = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[4]/ul/li[2]/form/div[1]/pre[1]";
                string elementjson = "/html/body/pre";
                string jsonText = ActiononWeb(By.XPath(btnGetAllEngineandTrans), By.XPath(btnTryIt), By.XPath(elementinputYear), By.XPath(elementState), By.XPath(elementCategory), By.XPath(elementMakeNicename), By.XPath(elementModelNicename), By.XPath(""), By.XPath(elementRequestURL), By.XPath(elementjson), true, true, true, year, "used", "", Makenicename, Modelnicename, "");
                
                if(jsonText != null)
                {
                    //objects is all make
                    var objects = JArray.Parse(jsonText);
                    int indexofEngine = 0;

                    foreach (JObject root in objects)
                    {
                        foreach (KeyValuePair<String, JToken> properties in root)
                        {
                            if (properties.Key.Contains("Count"))
                            {
                                break;
                            }
                            arrayAllEngineTrans = new ObjectEngineTrans[properties.Value.Count()];
                            foreach (JObject memPro in properties.Value)
                            {
                                if (indexofEngine >= arrayAllEngineTrans.Count())
                                {
                                    break;
                                }

                                arrayAllEngineTrans[indexofEngine] = new ObjectEngineTrans();
                                //root is make with id, name, nicename and models[]
                                foreach (KeyValuePair<String, JToken> app in memPro)
                                {
                                    //app is member of root with 2 value: Key and value of Key
                                    if (app.Key.Contains("engine"))
                                    {
                                        string cylinder = "";
                                        string size = "";
                                        string config = "";

                                        foreach (JProperty engine in app.Value)
                                        {

                                            switch (engine.Name.ToString())
                                            {
                                                case "cylinder":
                                                    cylinder = engine.Value.ToString();
                                                    break;
                                                case "size":
                                                    size = engine.Value.ToString();
                                                    break;
                                                case "configuration":
                                                    config = engine.Value.ToString();
                                                    string textEngine = config + cylinder + ", " + size + "L";
                                                    arrayAllEngineTrans[indexofEngine].Engine = textEngine;
                                                    break;
                                                default:
                                                    break;
                                            }

                                        }
                                    }

                                    if (app.Key.Contains("transmission"))
                                    {
                                        foreach (JProperty trans in app.Value)
                                        {

                                            if (trans.Name.Contains("transmissionType"))
                                            {
                                                arrayAllEngineTrans[indexofEngine].Transmission = trans.Value.ToString();

                                            }

                                        }
                                        indexofEngine++;
                                    }
                                }
                            }
                        }
                    }
                }
                
                return true;
            }
            else
            {
                return false;
            }

        }

        public static void GetModelYearIdbyMakeandModelNicename(string APIKey, string Makenicename, string Modelnicename)
        {
            //navigate to URL
            PropertiesCollection.NavigatetoURL("http://edmunds.mashery.com/io-docs");
            //enter API key
            IWebElement elementAPIKey = PropertiesCollection.driver.FindElement(By.Id("apiKey"));
            CommonMethods.DoWait(1000);
            elementAPIKey.SendKeys(APIKey);
            //Click button Get 

            string btnGetAllModelYearId = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/div/span[1]";

            string elementState = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/form/table/tbody/tr[3]/td[2]/select";
            string elementCategory = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/form/table/tbody/tr[5]/td[2]/select";
            string elementMakeNicename = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/form/table/tbody/tr[1]/td[2]/input";
            string elementModelNicename = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/form/table/tbody/tr[2]/td[2]/input";
            string btnTryIt = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/form/input[5]";
            string elementRequestURL = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[3]/ul/li[1]/form/div[1]/pre[1]";
            string elementjson = "/html/body/pre";
            string jsonText = ActiononWeb(By.XPath(btnGetAllModelYearId), By.XPath(btnTryIt), By.XPath(""), By.XPath(elementState), By.XPath(elementCategory), By.XPath(elementMakeNicename), By.XPath(elementModelNicename), By.XPath(""), By.XPath(elementRequestURL), By.XPath(elementjson), false, true, true, "", "used", "", Makenicename, Modelnicename, "");

            //objects is all make
            var objects = JArray.Parse(jsonText);
            int indexofModelYearId = 0;


            foreach (JObject root in objects)
            {
                foreach (KeyValuePair<String, JToken> properties in root)
                {
                    if (properties.Key.Contains("Count"))
                    {
                        break;
                    }
                    arrayAllModelYearId = new ObjectModelYearId[properties.Value.Count()];
                    foreach (JObject memPro in properties.Value)
                    {
                        if (indexofModelYearId > arrayAllModelYearId.Count())
                        {
                            break;
                        }

                        //root is make with id, name, nicename and models[]
                        foreach (KeyValuePair<String, JToken> app in memPro)
                        {
                            //app is member of root with 2 value: Key and value of Key
                            switch (app.Key.ToString())
                            {
                                case "id":
                                    //write ModelYearId to array
                                    indexofModelYearId++;
                                    arrayAllModelYearId[indexofModelYearId - 1] = new ObjectModelYearId();
                                    arrayAllModelYearId[indexofModelYearId - 1].ModelYearId = app.Value.ToString();
                                    break;
                                case "styles":
                                    foreach (JObject styles in app.Value)
                                    {
                                        foreach (KeyValuePair<String, JToken> memStyle in styles)
                                        {
                                            switch (memStyle.Key.ToString())
                                            {
                                                case "name":
                                                    //write engine in Database sheet
                                                    arrayAllModelYearId[indexofModelYearId - 1].Engine = memStyle.Value.ToString();
                                                    break;
                                                case "trim":
                                                    //write trim in Database sheet
                                                    arrayAllModelYearId[indexofModelYearId - 1].Trim = memStyle.Value.ToString();
                                                    break;
                                                case "submodel":
                                                    foreach (JObject submodel in app.Value)
                                                    {
                                                        foreach (KeyValuePair<String, JToken> memsubModel in submodel)
                                                        {
                                                            //write body in Database sheet
                                                            if (memsubModel.Key.Contains("body"))
                                                            {
                                                                arrayAllModelYearId[indexofModelYearId - 1].Body = memStyle.Value.ToString();
                                                            }
                                                            break;
                                                        }
                                                    }
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        public static bool GetScheduleMantenancebyModelYearId(string APIKey, string ModelYearId)
        {
            //navigate to URL
            PropertiesCollection.NavigatetoURL("http://edmunds.mashery.com/io-docs");

            if (CommonMethods.IsElementPresent(By.Id("apiKey")))
            {
                //enter API key
                IWebElement elementAPIKey = PropertiesCollection.driver.FindElement(By.Id("apiKey"));
                //new WebDriverWait(PropertiesCollection.driver, 20).until(ExpectedConditions.visibilityOfElementLocated(By.Id("apiKey")));

                CommonMethods.DoWait(1000);

                elementAPIKey.SendKeys(APIKey);
                //Click button Get 

                string btnGetAllSM = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[15]/ul/li[2]/div/span[1]";
                string elementmodelyearId = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[15]/ul/li[2]/form/table/tbody/tr[1]/td[2]/input";
                string btnTryIt = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[15]/ul/li[2]/form/input[5]";
                string elementRequestURL = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[15]/ul/li[2]/form/div[1]/pre[1]";
                string elementjson = "/html/body/pre";
                string jsonText = ActiononWeb(By.XPath(btnGetAllSM), By.XPath(btnTryIt), By.XPath(""), By.XPath(""), By.XPath(""), By.XPath(""), By.XPath(""), By.XPath(elementmodelyearId), By.XPath(elementRequestURL), By.XPath(elementjson), false, false, false, "", "", "", "", "", ModelYearId);

                if(jsonText != null)
                {
                    //objects is all make
                    var objects = JArray.Parse(jsonText);
                    int indexofSM = 0;


                    foreach (JObject root in objects)
                    {
                        foreach (KeyValuePair<String, JToken> properties in root)
                        {
                            if (properties.Key.Contains("Count"))
                            {
                                break;
                            }
                            arrayAllSM = new ObjectSM[properties.Value.Count()];
                            foreach (JObject memPro in properties.Value)
                            {
                                if (indexofSM > arrayAllSM.Count())
                                {
                                    break;
                                }

                                //root is make with id, name, nicename and models[]
                                foreach (KeyValuePair<String, JToken> app in memPro)
                                {
                                    //app is member of root with 2 value: Key and value of Key
                                    switch (app.Key.ToString())
                                    {
                                        case "id":
                                            indexofSM++;
                                            arrayAllSM[indexofSM - 1] = new ObjectSM();
                                            break;
                                        case "engineCode":
                                            arrayAllSM[indexofSM - 1].EngineCode = app.Value.ToString();
                                            break;
                                        case "intervalMileage":
                                            arrayAllSM[indexofSM - 1].IntervalMileage = app.Value.ToString();
                                            break;
                                        case "intervalMonth":
                                            arrayAllSM[indexofSM - 1].IntervalMonth = app.Value.ToString();
                                            break;
                                        case "frequency":
                                            arrayAllSM[indexofSM - 1].Frequency = app.Value.ToString();
                                            break;

                                        case "action":
                                            arrayAllSM[indexofSM - 1].Action = app.Value.ToString();
                                            break;
                                        case "item":
                                            arrayAllSM[indexofSM - 1].Item = app.Value.ToString();
                                            break;
                                        case "itemDescription":
                                            arrayAllSM[indexofSM - 1].ItemDes = app.Value.ToString();
                                            break;
                                        case "laborUnits":
                                            arrayAllSM[indexofSM - 1].LaborUnits = app.Value.ToString();
                                            break;
                                        case "partUnits":
                                            arrayAllSM[indexofSM - 1].PartUnits = app.Value.ToString();
                                            break;
                                        case "driveType":
                                            arrayAllSM[indexofSM - 1].DriveType = app.Value.ToString();
                                            break;
                                        case "partCostPerUnit":
                                            arrayAllSM[indexofSM - 1].PartCostPerUnit = app.Value.ToString();
                                            break;


                                        default:
                                            break;

                                    }

                                }
                            }
                        }

                    }
                }
                
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}

