using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using GetDataOnWeb_v01;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using OpenQA.Selenium.Support.UI;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data;

namespace KeyProgrammerProcedure_v01
{
    public class SM_Ford
    {
        public static void WriteDataExcel()
        {
            string path = Environment.CurrentDirectory;
            string namefile = "SM2.xlsx";
            string foldername = Path.Combine(path, namefile);

            System.Threading.Thread.Sleep(500);

            int rowYMME = 2;
            int colYMME = 8;
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                workbook.Worksheets.Add("Database");
                var MySheet = workbook.Worksheets[1];
                //workbook.Worksheets.Add("Database");
                //var MySheet2 = workbook.Worksheets[2];

                //write data to excel for first page of Make
                WriteHeader(foldername, MySheet, rowYMME, colYMME, package);
                rowYMME = GetDataAll(MySheet, rowYMME, colYMME, package);
            }
        }
        public static void WriteHeader(string foldername, ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            //create header of YMME sheet
            MySheet.Cells[1, 1].Value = "Year";
            MySheet.Cells[1, 2].Value = "Make";
            MySheet.Cells[1, 3].Value = "Model";

            package.SaveAs(new System.IO.FileInfo(foldername));
        }
            
        public static int GetDataAll(ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            //select years on DDL
            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("flsYearSelect"));
            var listYears = elementYears.AsDropDown().Options;

            List<string> AllYears = CommonMethods.ListStringfromIList(listYears);
            int row = 2;
            foreach (string _listYear in AllYears)
            {
                if (String.IsNullOrEmpty(_listYear) || String.IsNullOrWhiteSpace(_listYear) || _listYear.Contains("Select") || _listYear.Contains("2018") || _listYear.Contains("2017"))
                {
                    continue;
                }
                else
                {
                    System.Threading.Thread.Sleep(1000);
                    elementYears.AsDropDown().SelectByText(_listYear);
                    MySheet.Cells[row,1].Value = _listYear ;
                    package.Save();

                    System.Threading.Thread.Sleep(1000);

                    IWebElement elementMakes = PropertiesCollection.driver.FindElement(By.Id("flsMakeSelect"));
                    var listMakes = elementMakes.AsDropDown().Options;

                    //get all Makes
                    List<string> AllMakes = CommonMethods.ListStringfromIList(listMakes);

                    //elementMakes.Click();
                    System.Threading.Thread.Sleep(2000);

                    foreach (string _listMake in AllMakes)
                    {
                        if (String.IsNullOrEmpty(_listMake) || String.IsNullOrWhiteSpace(_listMake) || _listMake.Contains("Select") || _listYear.Contains("Mercury"))
                        {
                            continue;
                        }
                        else
                        {

                            System.Threading.Thread.Sleep(2000);
                            elementMakes.AsDropDown().SelectByText(_listMake);
                            MySheet.Cells[row, 2].Value = _listMake;
                            package.Save();

                            IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("flsModelSelect"));
                            
                            //IList<IWebElement> listoptionModel = elementModel.FindElements(By.TagName("option"));

                            var listModel = elementModel.AsDropDown().Options;
                            //get all Model
                            List<string> AllModel = CommonMethods.ListStringfromIList(listModel);

                            //Get Option Model 
                            IWebElement FindOpModel = PropertiesCollection.driver.FindElement(By.CssSelector("div.fls-form-section:nth-child(3) > select:nth-child(2) > option:nth-child(29)"));
                            FindOpModel.Click();

                            foreach (string _listModel in AllModel)
                            {
                                if (String.IsNullOrEmpty(_listModel) || String.IsNullOrWhiteSpace(_listModel) || _listModel.Contains("Select") || _listModel.Contains("C-MAX Energi") || _listModel.Contains("C-MAX Hybrid") || _listModel.Contains("E-350") || _listModel.Contains("E-450") || _listModel.Contains("Edge") || _listModel.Contains("Escape")|| _listModel.Contains("Expedition") || _listModel.Contains("Explorer") || _listModel.Contains("F-150")|| _listModel.Contains("F-250") || _listModel.Contains("F-350") || _listModel.Contains("F-450") || _listModel.Contains("F-550") || _listModel.Contains("F-650") || _listModel.Contains("F-750") || _listModel.Contains("Fiesta") || _listModel.Contains("Flex"))//|| _listModel.Contains("GT") || _listModel.Contains("Mustang") || _listModel.Contains("Taurus") || _listModel.Contains("Transit")) //|| _listModel.Contains("Focus")
                                {
                                    continue;
                                  
                                }

                                else
                                {
                                    System.Threading.Thread.Sleep(3000);
                                    
                                    //elementModel.AsDropDown().SelectByText(_listModel);

                                    MySheet.Cells[row, 3].Value = _listModel;
                                    package.Save();                                  

                                    //Click button
                                    IWebElement btnFindVehicle = PropertiesCollection.driver.FindElement(By.ClassName("fls-search-form-ymm-actions"));
                                    btnFindVehicle.Click();
                                    System.Threading.Thread.Sleep(4000);
                                    
                                    //Send text 
                                    IWebElement FindInput1 = PropertiesCollection.driver.FindElement(By.Name("currentMileage"));
                                    FindInput1.SendKeys("3000");                                    

                                    //Send text 
                                    IWebElement FindInput2 = PropertiesCollection.driver.FindElement(By.Name("averageMileage"));
                                    FindInput2.SendKeys("30");


                                    //IWebElement FindButton1 = PropertiesCollection.driver.FindElement(By.CssSelector("label.ng-scope:nth-child(3) > span:nth-child(2)"));
                                    //FindButton1.Click();


                                    //Click Driving Condition
                                    IWebElement elementCondition = PropertiesCollection.driver.FindElement(By.TagName("select"));
                                    var listCondition = elementCondition.AsDropDown().Options;
                                    List<string> AllCondition = CommonMethods.ListStringfromIList(listCondition);

                                    foreach (string _listCondition in AllCondition)
                                    {
                                        if (String.IsNullOrEmpty(_listCondition) || String.IsNullOrWhiteSpace(_listCondition) || _listCondition.Contains("Select") || _listCondition.Contains("Extensive idling and/or driving at low speeds") || _listCondition.Contains("Operating in Dusty Conditions"))
                                        {
                                            continue;

                                        }
                                        else
                                        {
                                            System.Threading.Thread.Sleep(1500);
                                            elementCondition.AsDropDown().SelectByText(_listCondition);
                                            MySheet.Cells[row, 4].Value = _listCondition;
                                            package.Save();

                                            IWebElement FindButton2 = PropertiesCollection.driver.FindElement(By.CssSelector("#ng-app-SYNC > div.fls-content.clearer > div.ng-scope > div.vehicle-details.ng-scope > div > div.fls-maintenance-schedule-vehicle-details.fls-inner-block > div > form > div:nth-child(4) > div > label:nth-child(1) > span"));
                                            FindButton2.Click();

                                            System.Threading.Thread.Sleep(1500);

                                            //IWebElement FindButton3 = PropertiesCollection.driver.FindElement(By.CssSelector("#ng-app-SYNC > div.fls-content.clearer > div.ng-scope > div.vehicle-details.ng-scope > div > div.fls-maintenance-schedule-vehicle-details.fls-inner-block > div > form > div:nth-child(5) > div > label:nth-child(2) > span"));
                                            //FindButton3.Click();
                                            //System.Threading.Thread.Sleep(1500);

                                            //IWebElement FindButton4 = PropertiesCollection.driver.FindElement(By.CssSelector("#ng-app-SYNC > div.fls-content.clearer > div.ng-scope > div.vehicle-details.ng-scope > div > div.fls-maintenance-schedule-vehicle-details.fls-inner-block > div > form > div:nth-child(5) > div > label:nth-child(2) > span"));
                                            //FindButton4.Click();
                                            //System.Threading.Thread.Sleep(1500);

                                            //Click button                                        
                                            IWebElement btnUpdateVehicle = PropertiesCollection.driver.FindElement(By.CssSelector("#ng-app-SYNC > div.fls-content.clearer > div.ng-scope > div.vehicle-details.ng-scope > div > div.fls-maintenance-schedule-vehicle-details.fls-inner-block > div > form > div.fls-update-details-cta > input"));                                                                                          
                                            System.Threading.Thread.Sleep(1500);
                                            btnUpdateVehicle.Click();
                                            System.Threading.Thread.Sleep(3000);

                                            //Get Interval  & service name
                                            IWebElement eleIntervalData = PropertiesCollection.driver.FindElement(By.ClassName("fls-slider-mileage"));
                                            IList<IWebElement> listInterval = eleIntervalData.FindElements(By.ClassName("ng-scope"));
                                            List<string> AllIntervalData = CommonMethods.ListStringfromIList(listInterval);
                                            int _interval = 0;
                                            int col = 5;
                                            foreach (string IntervalItem in AllIntervalData)
                                            {
                                                if (listInterval.Count != 0)
                                                {
                                                    listInterval[_interval].Click();
                                                    MySheet.Cells[row, col].Value = IntervalItem;
                                                    package.Save();


                                                    _interval++;

                                                    System.Threading.Thread.Sleep(1000);
                                                    IWebElement elementServ = PropertiesCollection.driver.FindElement(By.ClassName("fls-checkup-content-container"));

                                                    IList<IWebElement> listServ = elementServ.FindElements(By.ClassName("ng-binding"));
                                                    IList<IWebElement> listService = elementServ.FindElements(By.ClassName("ul"));
                                                    List<string> AllServiceName = CommonMethods.ListStringfromIList(listServ);

                                                      //foreach(string ServiceItem in AllServiceName)
                                                      //{
                                                        System.Threading.Thread.Sleep(1000);
                                                        MySheet.Cells[row+1, col].Value = elementServ.Text;
                                                        package.Save();                                                       
                                                      //}                                                   
                                                    col++;
                                                  }
                                                 else { continue; }                                                                                                                                                                                            
                                            }

                                            IWebElement next = PropertiesCollection.driver.FindElement(By.ClassName("fls-slider-next"));
                                            next.Click();

                                            IWebElement eleInterval2 = PropertiesCollection.driver.FindElement(By.ClassName("fls-slider-mileage"));
                                            IList<IWebElement> listInterval2 = eleInterval2.FindElements(By.ClassName("ng-scope"));
                                            List<string> AllIntervalData2 = CommonMethods.ListStringfromIList(listInterval2);

                                            int _interval2 = 0;
                                            foreach (string IntervalItem2 in AllIntervalData2)
                                            {
                                                if (listInterval2.Count != 0)
                                                {
                                                    listInterval2[_interval2].Click();
                                                    MySheet.Cells[row, col].Value = IntervalItem2;
                                                    package.Save();
                                                    _interval2++;                                                    
                                                    IWebElement elementServ2 = PropertiesCollection.driver.FindElement(By.ClassName("fls-checkup-content-container"));

                                                    IList<IWebElement> listServ2 = elementServ2.FindElements(By.ClassName("ng-binding"));
                                                    IList<IWebElement> listService2 = elementServ2.FindElements(By.ClassName("ul"));
                                                    List<string> AllServiceName2 = CommonMethods.ListStringfromIList(listServ2);

                                                    //  foreach(string ServiceItem in AllServiceName)
                                                    // {
                                                  
                                                    MySheet.Cells[row+1, col].Value = elementServ2.Text;
                                                    package.Save();

                                                    if (IntervalItem2 == "null")
                                                    {
                                                        continue;
                                                    }
                                                    col++;
                                                    //  }
                                                }
                                           
                                               // col++;
                                            }
                                            continue;
                                            //IWebElement FindArrow = PropertiesCollection.driver.FindElement(By.ClassName("fls-vehicle-select-handle"));
                                            //FindArrow.Click();
                                            //System.Threading.Thread.Sleep(1000);

                                            //IWebElement FindSelectCar = PropertiesCollection.driver.FindElement(By.ClassName("fls-vehicle-select-another span-link ng-binding"));
                                            //FindSelectCar.Click();
                                            //System.Threading.Thread.Sleep(1500);
                                            //string url = "https://owner.ford.com/tools/account/maintenance/maintenance-schedule.html#/details";
                                            //PropertiesCollection.driver.Navigate().GoToUrl(url);
                                            //Button Next Click                                                                               
                                        }                                       
                                    }
                                }

                            }
                        }
                    }
                    row++;
                }              
            }
            return rowYMME;
        }
    }
}
