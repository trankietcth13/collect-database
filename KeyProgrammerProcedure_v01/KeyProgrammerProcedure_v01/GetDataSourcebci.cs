using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OpenQA.Selenium;
using GetDataOnWeb_v01;
using Excel = Microsoft.Office.Interop.Excel;

using System.IO;

namespace KeyProgrammerProcedure_v01
{
    class GetDataSourcebci
    {
        public static List<string> listMake = new List<string>
        {
            "ACURA", //done
            "LINCOLN",
            "MAZDA",
            "MERCURY",
            "ACURA", //done
            "AUDI", //done
            "BMW", //done
            "BUICK", //done
            "CADILLAC", //done
            "CHEVROLET", //done
            "CHRYSLER", //done
            "Daewoo",
            "DODGE",
            "EAGLE",
            "GEO",
            "GMC",
            "HONDA",
            "HUMMER",
            "HYUNDAI",
            "INFINITI",
            "ISUZU",
            "JAGUAR",
            "JEEP",
            "KIA",
            "LAND_ROVER",
            "LEXUS",
            "MERCEDES",
            "MINI",
            "MITSUBISHI",
            "NISSAN",
            "OLDSMOBILE",
            "PLYMOUTH",
            "PONTIAC",
            "PORSCHE",
            "SAAB",
            "SATURN",
            "SCION",
            "SMART",
            "SUBARU",
            "SUZUKI",
            "TOYOTA",
            "VOLKSWAGEN",
            "VOLVO"
        };
        public static void WriteDataExcel()
        {
            string path = Environment.CurrentDirectory;
            string namefile = "SourceBCIData.xlsx";
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
                rowYMME = GetdataBCI(MySheet, rowYMME, colYMME, package);
            }
        }
        public static void WriteHeader(string foldername, ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            //create header of YMME sheet
            MySheet.Cells[1, 1].Value = "Make";
            MySheet.Cells[1, 2].Value = "Year";
            MySheet.Cells[1, 3].Value = "Model";
            MySheet.Cells[1, 4].Value = "Engine";
            MySheet.Cells[1, 5].Value = "BCI Group No.";
            MySheet.Cells[1, 6].Value = "CCA";

            package.SaveAs(new System.IO.FileInfo(foldername));
        }

        public static int GetdataBCI(ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            //select Makes on DDL
            IWebElement elementMakes = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
            var listMakes = elementMakes.AsDropDown().Options;
            //get all Makes
            List<string> AllMakes = CommonMethods.ListStringfromIList(listMakes);
            int row = 2;
            foreach(string _listMakes in AllMakes)
            {
                //skip Please choose Make
                if (String.IsNullOrEmpty(_listMakes) || String.IsNullOrWhiteSpace(_listMakes) || _listMakes.Contains("-select-") || _listMakes.Contains("Cadillac") || _listMakes.Contains("Buick") || _listMakes.Contains("BMW") || _listMakes.Contains("Acura") || _listMakes.Contains("Audi") || _listMakes.Contains("Chevrolet") || _listMakes.Contains("Chrysler"))//Chrysler
                {
                    continue;
                }
                else
                {
                    //Compare _listmake with listmake had created.
                    foreach (string MakeOnList in GetDataSourcebci.listMake)
                    {
                        if (_listMakes.ToUpper().Contains(MakeOnList))
                        {
                            System.Threading.Thread.Sleep(3000);

                            IWebElement elementMakes4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
                            //select Make on DDL
                            elementMakes4.AsDropDown().SelectByText(_listMakes);

                            MySheet.Cells[row, 1].Value = _listMakes;
                            package.Save();
                            System.Threading.Thread.Sleep(4000);

                            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                            var listYears = elementYears.AsDropDown().Options;

                            //get all Models
                            List<string> AllYears = CommonMethods.ListStringfromIList(listYears);

                            foreach (string _listYears in AllYears)
                            {
                                //skip Please choose Make
                                if (String.IsNullOrEmpty(_listYears) || String.IsNullOrWhiteSpace(_listYears) || _listYears.Contains("-select-")|| _listYears.Contains("2015") || _listYears.Contains("2016"))//|| _listYears.Contains("2013") || _listYears.Contains("2014") || _listYears.Contains("2012") || _listYears.Contains("2011")) //|| _listYears.Contains("2010") || _listYears.Contains("2009") || _listYears.Contains("2008")) // || _listYears.Contains("2009")|| _listYears.Contains("2015") || _listYears.Contains("2014") || _listYears.Contains("2013") || _listYears.Contains("2012") || _listYears.Contains("2011") || _listYears.Contains("2010") || _listYears.Contains("2009")) //|| _listYears.Contains("2013")) //|| _listYears.Contains("2016")
                                {
                                    continue;
                                }
                                int year = Convert.ToInt32(_listYears);
                                if (year < 2008 || year > 2016)
                                {
                                    continue;
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(3000);

                                    IWebElement elemenYears4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                                    //select Make on DDL
                                    elemenYears4.AsDropDown().SelectByText(_listYears);

                                    MySheet.Cells[row, 2].Value = _listYears;
                                    package.Save();

                                    System.Threading.Thread.Sleep(5000);

                                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                    var listModels = elementModel.AsDropDown().Options;

                                    //get all Models
                                    List<string> AllModels = CommonMethods.ListStringfromIList(listModels);
                                    foreach (string _listModels in AllModels)
                                    {
                                        //skip Please choose Model
                                        if (String.IsNullOrEmpty(_listModels) || String.IsNullOrWhiteSpace(_listModels) || _listModels.Contains("-select-"))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            System.Threading.Thread.Sleep(1000);
                                            IWebElement elementModel4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                            //select Model on DDL
                                            elementModel4.AsDropDown().SelectByText(_listModels);
                                            MySheet.Cells[row, 3].Value = _listModels;
                                            package.Save();
                                            System.Threading.Thread.Sleep(5000);                                                                                   

                                            #region Get BCI & CCA
                                            IList<IWebElement> element_table = PropertiesCollection.driver.FindElements(By.ClassName("subgridview"));
                                            foreach (var itemOfTable in element_table)
                                            {                                                
                                                //Get data
                                                IWebElement FindTagTd = PropertiesCollection.driver.FindElement(By.ClassName("gridview"));
                                                IList<IWebElement> TagTd_table = FindTagTd.FindElements(By.TagName("td"));                                                
                                                List<string> AlltagTd = CommonMethods.ListStringfromIList(TagTd_table);

                                                int rowofTagTd = 0;
                                                int colofTagTd = 0;

                                                                                            
                                                foreach (var item_TagTd in TagTd_table)
                                                {
                                                    if (_listModels.Contains("Challenger") || _listModels.Contains("Charger") || AlltagTd == null || AlltagTd.Contains(""))
                                                    {
                                                        IWebElement findPlusButton = PropertiesCollection.driver.FindElement(By.Id("MainContent_gvApplication_btnDetail_0"));
                                                        findPlusButton.Click();
                                                        System.Threading.Thread.Sleep(1000);

                                                        IWebElement findPlusButton1 = PropertiesCollection.driver.FindElement(By.Id("MainContent_gvApplication_btnDetail_1"));
                                                        findPlusButton1.Click();
                                                        System.Threading.Thread.Sleep(1000);

                                                        IWebElement findPlusButton2 = PropertiesCollection.driver.FindElement(By.Id("MainContent_gvApplication_btnDetail_2"));
                                                        findPlusButton2.Click();
                                                        System.Threading.Thread.Sleep(1000);

                                                        IWebElement findPlusButton3 = PropertiesCollection.driver.FindElement(By.Id("MainContent_gvApplication_btnDetail_3"));
                                                        findPlusButton3.Click();
                                                        System.Threading.Thread.Sleep(1000);
                                                    }

                                                    else if (_listMakes.Contains("BMW") || _listMakes == "BMW")
                                                            {
                                                                rowofTagTd++;
                                                                colofTagTd++;
                                                                switch (rowofTagTd)
                                                                {
                                                                    //case 2:
                                                                    //    //write BCI Group
                                                                    //    // Makes
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;

                                                                    //case 3:
                                                                    //    //write CCA 
                                                                    //    //Models
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;
                                                                    case 5:
                                                                         //Engines                                                                       
                                                                         MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                         package.Save();                                                                                                                                                                                                                           
                                                                         break;
                                                            //case 6:
                                                            //    //write CCA 2
                                                            //    //Engines
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;

                                                            //case 7:                                                            
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;
                                                            case 8:  //BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 9: //CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 13:  //sub_BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 14: //sub_CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    //case 16:                                                            
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 18:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 19:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                            case 21:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 20:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 23:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 24:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                            case 25:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 27:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 28:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 29:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 30:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 31:
                                                            //            MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //            package.Save();
                                                            //            break;

                                                            case 33:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 34:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 37:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 40:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 41:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 45:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 46:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 50:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 51:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 55:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 56:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 60:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 61:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 65:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 66:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            default:
                                                                break;
                                                                }
                                                            }

                                                    else if(_listMakes.Contains("Acura") || _listMakes == "Acura")
                                                            {
                                                                rowofTagTd++;
                                                                colofTagTd++;
                                                                switch (rowofTagTd)
                                                                {
                                                                    //case 2:
                                                                    //    //write BCI Group
                                                                    //    // Makes
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;

                                                                    //case 3:
                                                                    //    //write CCA 
                                                                    //    //Models
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;
                                                                    case 5:
                                                                        //Engines                                                                       
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                    //case 6:
                                                                    //    //write CCA 2
                                                                    //    //Engines
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;

                                                                    //case 7:                                                            
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;
                                                                    case 8:  //BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 9: //CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 13:  //sub_BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 14: //sub_CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                            case 16:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 18:
                                                            //            MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //            package.Save();
                                                            //            break;

                                                                    case 19:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                            case 20:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                            //case 21:
                                                            //            MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //            package.Save();
                                                            //            break;

                                                            //case 20:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 23:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 24:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 25:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 27:  //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 28:
                                                            //            MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //            package.Save();
                                                            //            break;

                                                                    //case 29:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 30:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                            case 31:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 33:
                                                            //            MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //            package.Save();
                                                            //            break;

                                                                    case 34:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 37:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 40:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 41:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 45:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 46:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 50:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 51:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 55:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 56:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 60:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 61:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 65:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 66:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                    default:
                                                                        break;
                                                                }
                                                            }

                                                    else if (_listMakes.Contains("Audi") || _listMakes == "Audi")
                                                        {
                                                            rowofTagTd++;
                                                            colofTagTd++;
                                                            switch (rowofTagTd)
                                                            {
                                                            //case 2:
                                                            //    //write BCI Group
                                                            //    // Makes
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;

                                                            //case 3:
                                                            //    //write CCA 
                                                            //    //Models
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;
                                                            case 5:
                                                                //Engines                                                                       
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            //case 6:
                                                            //    //write CCA 2
                                                            //    //Engines
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;

                                                            //case 7:                                                            
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;
                                                            case 8:  //BCI                                                            
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 9: //CCA
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 13:  //sub_BCI                                                            
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 14: //sub_CCA
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 16:                                                            
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 18:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 19:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            //case 21:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 20:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 23:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 24:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 25:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 27:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 28:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 29:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 30:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 31:
                                                            //            MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //            package.Save();
                                                            //            break;

                                                            case 33:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 34:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 38:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 39:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 41:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 46: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 49:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 50:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 51:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 54:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 55:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 59:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 60:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 64:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 65:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 69:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 70:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 74:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 75:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 79:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 80:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            default:
                                                                break;
                                                        }
                                                    }

                                                    else if (_listMakes.Contains("Buick") || _listMakes == "Buick")
                                                            {
                                                            rowofTagTd++;
                                                            colofTagTd++;
                                                            switch (rowofTagTd)
                                                            {
                                                            //case 2:
                                                            //    //write BCI Group
                                                            //    // Makes
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;

                                                            //case 3:
                                                            //    //write CCA 
                                                            //    //Models
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;
                                                            case 5:
                                                                //Engines                                                                       
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            //case 6:
                                                            //    //write CCA 2
                                                            //    //Engines
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;

                                                            //case 7:                                                            
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;
                                                            case 8:  //BCI                                                            
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 9: //CCA
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 13:  //sub_BCI                                                            
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 14: //sub_CCA
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 16: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 18:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 19:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            //case 21:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 20:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 23:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 24:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 25:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 27: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 28:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 29:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 30:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 31:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 33:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 34:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 38:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 39:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 41:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 46: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 49:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 50:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 51:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 54:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 55:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 59:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 60:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 64:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 65:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 69:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 70:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 74:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 75:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 79:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 80:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            default:
                                                                break;
                                                        }
                                                    }

                                                    else if (_listMakes.Contains("Cadillac") || _listMakes == "Cadillac")
                                                        {
                                                        rowofTagTd++;
                                                        colofTagTd++;
                                                        switch (rowofTagTd)
                                                        {
                                                           
                                                            case 5: //Engines           

                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            //case 6:
                                                            //    //write CCA 2
                                                            //    //Engines
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;

                                                            //case 7:                                                            
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    break;
                                                            case 8:  //BCI                                                            
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 9: //CCA
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 13:  //sub_BCI                                                            
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 14: //sub_CCA
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 16: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 19:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;
                                                            //case 21:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 20:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 21: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            //case 23:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 24:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 25:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 27: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 28:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 29:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 30:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 31:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 33:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 34:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 37: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            //case 38:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            //case 39:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 40:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 41:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 45:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 46: //Get Engine
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            //case 49:
                                                            //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                            //    package.Save();
                                                            //    break;

                                                            case 50:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 51:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 54:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 55:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 59:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 60:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 64:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 65:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;


                                                            case 69:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 70:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 74:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 75:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 79:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            case 80:
                                                                MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                package.Save();
                                                                break;

                                                            default:
                                                                break;
                                                        }
                                                    }

                                                    else if (_listMakes.Contains("Chevrolet") || _listMakes == "Chevrolet")
                                                            {
                                                            rowofTagTd++;
                                                            colofTagTd++;                                                               
                                                                
                                                            switch (rowofTagTd)
                                                            {

                                                                case 5: //Engines           
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;
                                                                //case 6:
                                                                //    //write CCA 2
                                                                //    //Engines
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    break;

                                                                //case 7:                                                            
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    break;
                                                                case 8:  //BCI                                                            
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 9: //CCA
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 13:  //sub_BCI                                                                    
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 14: //sub_CCA
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 16: //Get Engine
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 19:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;
                                                                //case 21:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                case 20:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 21: //Get Engine
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;


                                                                //case 23:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                case 24:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 25:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 27: //Get Engine
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                //case 28:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                //case 29:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                case 30:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 31:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                //case 33:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                //case 34:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                case 37: //Get Engine
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                //case 38:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                //case 39:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                case 40:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 41:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 45:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 46: //Get Engine
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;


                                                                //case 49:
                                                                //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                //    package.Save();
                                                                //    break;

                                                                case 50:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 51:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 54:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;


                                                                case 55:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 59:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;


                                                                case 60:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;


                                                                case 64:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 65:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;


                                                                case 69:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 70:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 74:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 75:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 79:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                case 80:
                                                                    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    package.Save();
                                                                    break;

                                                                default:
                                                                    break;
                                                            }
                                                            }

                                                    else if (_listMakes.Contains("Chrysler") || _listMakes == "Chrysler")
                                                            {
                                                                rowofTagTd++;
                                                                colofTagTd++;
                                                                switch (rowofTagTd)
                                                                {

                                                                    case 5: //Engines           

                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                    //case 6:
                                                                    //    //write CCA 2
                                                                    //    //Engines
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;

                                                                    //case 7:                                                            
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    break;
                                                                    case 8:  //BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 9: //CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 13:  //sub_BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 14: //sub_CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 16: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 19:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                    //case 21:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 20:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 21: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    //case 23:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 24:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 25:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 27: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    //case 28:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 29:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 30:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 31:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    //case 33:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    //case 34:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 37: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    //case 38:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    //case 39:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 40:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 41:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 45:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 46: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    //case 49:
                                                                    //    MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                    //    package.Save();
                                                                    //    break;

                                                                    case 50:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 51:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 54:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 55:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 59:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 60:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 64:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 65:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 69:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 70:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 74:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 75:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 79:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 80:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    default:
                                                                        break;
                                                                }
                                                            }

                                                    else if (_listMakes.Contains("Dodge") || _listMakes == "Dodge")
                                                    {
                                                           rowofTagTd++;
                                                           colofTagTd++;
                                                           switch (rowofTagTd)
                                                                {

                                                                    case 5: //Engines           

                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;                                                                    
                                                                    case 8:  //BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 9: //CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 13:  //sub_BCI                                                            
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 14: //sub_CCA
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 16: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 18:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                   case 19:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                  
                                                                    case 20:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 21: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                
                                                                    case 24:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 25:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 27: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;                                                                

                                                                    case 29:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 30:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 31:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                  
                                                                    case 37: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                
                                                                    case 40:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 41:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 48: //Get Engine
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                
                                                                    case 51:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 52:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 54:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 56:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 57:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;
                                                                  
                                                                    case 60:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 64:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 65:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;


                                                                    case 69:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 70:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 74:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 75:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 79:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    case 80:
                                                                        MySheet.Cells[row, colofTagTd].Value = item_TagTd.Text.ToString();
                                                                        package.Save();
                                                                        break;

                                                                    default:
                                                                        break;
                                                                }
                                                    }
                                                }
                                                row++;
                                                //TagTdEngine++;
                                                //System.Threading.Thread.Sleep(3000);                                               
                                            }
                                            #endregion
                                        }
                                    }                                   
                                }
                            }
                            break;                           
                        }                        
                    }
                }
            }
            return rowYMME;
        }
    }
}
