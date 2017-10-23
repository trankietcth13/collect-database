using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace GetDataOnWeb_v01
{
    public static class Sourcebci
    {
        public static List<string> listMake = new List<string>
        {
            "ACURA",
            "LINCOLN",
            "MAZDA",
            "MERCURY",
            "ACURA",
            "AUDI",
            "BMW",
            "BUICK",
            "CADILLAC",
            "CHEVROLET",
            "CHRYSLER",
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
        public static void GetDataAll()
        {
            Excel.Application MyApp = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook MyBook = MyApp.Workbooks.Add(misValue);
            Excel.Worksheet MySheet = (Excel.Worksheet)MyBook.Sheets[1];
            Excel.Range xlRange = MySheet.UsedRange;
            MyApp.Visible = false;

            //number of row and column in data file
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //create header 
            MySheet.Cells[1, 1] = "YEAR";
            MySheet.Cells[1, 2] = "MAKE";
            MySheet.Cells[1, 3] = "MODEL";
            MySheet.Cells[1, 4] = "ENGINE";

            MySheet.Cells[1, 5] = "BCI Group No.";
            MySheet.Cells[1, 6] = "CCA";
            //MySheet.Cells[1, 7] = "Amp Hour";
            //MySheet.Cells[1, 8] = "Notes";

            //row of data
            int row = 2;

            string path = Environment.CurrentDirectory;
            string foldername = Path.Combine(path, "outputSourceBCI.xlsx");

            //save excel
            MyBook.SaveAs(foldername);
            System.Threading.Thread.Sleep(500);

            //select Makes on DDL
            IWebElement elementMakes = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
            var listMakes = elementMakes.AsDropDown().Options;

            //get all Makes
            List<string> AllMakes = CommonMethods.ListStringfromIList(listMakes);

            foreach (string linkTextMake in AllMakes)
            {
                //skip Please choose Make
                if (String.IsNullOrEmpty(linkTextMake) || String.IsNullOrWhiteSpace(linkTextMake) || linkTextMake.Contains("-select-"))
                {
                    continue;
                }
                else
                {
                    //compare with listmake 
                    foreach (string yearonlist in Sourcebci.listMake)
                    {
                        if (linkTextMake.ToUpper().Contains(yearonlist))
                        {
                            System.Threading.Thread.Sleep(3000);

                            IWebElement elementMakes4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
                            //select Make on DDL
                            elementMakes4.AsDropDown().SelectByText(linkTextMake);

                            System.Threading.Thread.Sleep(3000);

                            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                            var listYears = elementYears.AsDropDown().Options;

                            //get all Models
                            List<string> AllYears = CommonMethods.ListStringfromIList(listYears);

                            foreach (string linkTextYear in AllYears)
                            {
                                //skip Please choose Make
                                if (String.IsNullOrEmpty(linkTextYear) || String.IsNullOrWhiteSpace(linkTextYear) || linkTextYear.Contains("-select-"))
                                {
                                    continue;
                                }
                                int year = Convert.ToInt32(linkTextYear);
                                if (year < 1996 || year > 2016)
                                {
                                    continue;
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(1000);

                                    IWebElement elemenYears4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                                    //select Make on DDL
                                    elemenYears4.AsDropDown().SelectByText(linkTextYear);

                                    System.Threading.Thread.Sleep(5000);

                                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                    var listModels = elementModel.AsDropDown().Options;

                                    //get all Models
                                    List<string> AllModels = CommonMethods.ListStringfromIList(listModels);

                                    foreach (string linkTextModel in AllModels)
                                    {
                                        //skip Please choose Model
                                        if (String.IsNullOrEmpty(linkTextModel) || String.IsNullOrWhiteSpace(linkTextModel) || linkTextModel.Contains("-select-"))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            System.Threading.Thread.Sleep(1000);
                                            IWebElement elementModel4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                            //select Model on DDL
                                            elementModel4.AsDropDown().SelectByText(linkTextModel);
                                            System.Threading.Thread.Sleep(5000);

                                            //click Find button 
                                            IWebElement FindButton = PropertiesCollection.driver.FindElement(By.Id("MainContent_btnSearch1"));

                                            FindButton.Click();
                                            System.Threading.Thread.Sleep(5000);

                                            //write year
                                            MySheet.Cells[row, 1] = linkTextYear;
                                            //write make
                                            MySheet.Cells[row, 2] = linkTextMake;
                                            //write model
                                            MySheet.Cells[row, 3] = linkTextModel;
                                            //write engine
                                            //MySheet.Cells[row, 4] = linkTextEngine; 

                                            IList<IWebElement> element_table = PropertiesCollection.driver.FindElements(By.ClassName("subgridview"));
                                            foreach(var itemOfTable in element_table)
                                            {
                                                IList<IWebElement> TagTd_table = itemOfTable.FindElements(By.TagName("td"));

                                                int rowofTagTd = 0;
                                                foreach (var item_TagTd in TagTd_table)
                                                {
                                                    rowofTagTd++;
                                                    switch (rowofTagTd)
                                                    {
                                                        case 2:
                                                            //write BCI Group
                                                            MySheet.Cells[row, 5] = item_TagTd.Text.ToString();
                                                            break;
                                                        case 3:
                                                            //write CCA 
                                                            MySheet.Cells[row, 6] = item_TagTd.Text.ToString();
                                                            break;
                                                        //case 4:
                                                        //    //write Amp Hour
                                                        //    MySheet.Cells[row, 7] = item_TagTd.Text.ToString();
                                                        //    break;
                                                        //case 5:
                                                        //    //write Notes
                                                        //    MySheet.Cells[row, 8] = item_TagTd.Text.ToString();
                                                        //    break;
                                                        default:
                                                            break;
                                                    }

                                                }
                                                row++;
                                                System.Threading.Thread.Sleep(1000);
                                            }                                                                                       
                                        }
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }

        }
    }
}
