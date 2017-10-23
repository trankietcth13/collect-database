using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.ObjectModel;

namespace GetDataOnWeb_v01
{
    public static class BatteryFinder
    {

        //year text from drop down list
        public static string year;

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

            MySheet.Cells[1, 5] = "Group Size";
            MySheet.Cells[1, 6] = "Minimum Cold Cranking Amps";
            MySheet.Cells[1, 7] = "Description ";

            //row of data
            int row = 2;
            
            string path = Environment.CurrentDirectory;
            string foldername = Path.Combine(path, "outputBatteryFinder.xlsx");

            MyBook.SaveAs(foldername);
            //waitOnPage(5)
            //waitForPageUntilElementIsVisible(By.LinkText(linkTextRangeYear), 5);
            System.Threading.Thread.Sleep(500);

            //select year on DDL
            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
            var listYears = elementYears.AsDropDown().Options;

            //get all Years
            List<string> AllYears = CommonMethods.ListStringfromIList(listYears);

            foreach (string linkTextYear in AllYears)
            {
                if(linkTextYear.Contains("1995"))
                {
                    break;
                }
                //skip Please choose Make
                if (String.IsNullOrEmpty(linkTextYear) || String.IsNullOrWhiteSpace(linkTextYear) || linkTextYear.Contains("Please Choose Year") || linkTextYear.Contains("2018") || linkTextYear.Contains("2017"))
                {
                    continue;
                }
                else
                {
                    System.Threading.Thread.Sleep(1000);

                    IWebElement elemenYears4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                    //select Make on DDL
                    elemenYears4.AsDropDown().SelectByText(linkTextYear);

                    System.Threading.Thread.Sleep(1000);

                    IWebElement elementMakes = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                    var listMakes = elementMakes.AsDropDown().Options;

                    //get all Makes
                    List<string> AllMakes = CommonMethods.ListStringfromIList(listMakes);

                    foreach (string linkTextMake in AllMakes)
                    {
                        //skip Please choose Make
                        if (String.IsNullOrEmpty(linkTextMake) || String.IsNullOrWhiteSpace(linkTextMake) || linkTextMake.Contains("Please Choose Make"))
                        {
                            continue;
                        }
                        else
                        {
                            System.Threading.Thread.Sleep(1000);
                            IWebElement elementMakes4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                            //select Make on DDL
                            elementMakes4.AsDropDown().SelectByText(linkTextMake);

                            System.Threading.Thread.Sleep(1000);

                            IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("product_finder_model"));
                            var listModels = elementModel.AsDropDown().Options;

                            //get all Models
                            List<string> AllModels = CommonMethods.ListStringfromIList(listModels);

                            foreach (string linkTextModel in AllModels)
                            {
                                //skip Please choose Model
                                if (String.IsNullOrEmpty(linkTextModel) || String.IsNullOrWhiteSpace(linkTextModel) || linkTextModel.Contains("Please Choose Model"))
                                {
                                    continue;
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(1000);
                                    IWebElement elementModel4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_model"));
                                    //select Model on DDL
                                    elementModel4.AsDropDown().SelectByText(linkTextModel);

                                    System.Threading.Thread.Sleep(1000);

                                    IWebElement elementEngine = PropertiesCollection.driver.FindElement(By.Id("product_finder_engine"));
                                    var listEngine = elementEngine.AsDropDown().Options;

                                    //get all Models
                                    List<string> AllEngine = CommonMethods.ListStringfromIList(listEngine);

                                    foreach (string linkTextEngine in AllEngine)
                                    {
                                        //skip Please choose Engine
                                        if (String.IsNullOrEmpty(linkTextEngine) || String.IsNullOrWhiteSpace(linkTextEngine) || linkTextEngine.Contains("Please Choose Engine"))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            System.Threading.Thread.Sleep(1000);
                                            //select Model on DDL
                                            IWebElement elementEngine4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_engine"));
                                            elementEngine4.AsDropDown().SelectByText(linkTextEngine);
                                            System.Threading.Thread.Sleep(1000);

                                            //click Find button 
                                            IWebElement FindButton = PropertiesCollection.driver.FindElement(By.Id("find_button"));

                                            FindButton.Click();

                                            System.Threading.Thread.Sleep(1000);

                                            //check Element exists
                                            //do if exist
                                            if (CommonMethods.IsElementPresent(By.ClassName("label")))
                                            {

                                                int tempcount = 0;
                                                int tempcountforXpath = 1;
                                                //get all value
                                                List<string> AllValues = CommonMethods.GetListfromtheSameElements(By.ClassName("value"));
                                                foreach (string value in AllValues)
                                                {
                                                    //value is Amps
                                                    if (tempcount % 2 != 0)
                                                    {
                                                        //write Amps
                                                        MySheet.Cells[row, 6] = value;

                                                        //do if exist
                                                        if (CommonMethods.IsElementPresent(By.XPath("/html/body/div/div/div/div[1]/section/div[3]/div[1]/div[2]")))
                                                        {

                                                            IWebElement elementDscrpt = PropertiesCollection.driver.FindElement(By.XPath("/html/body/div/div/div/div[1]/section/div[3]/div[" + tempcountforXpath + "]/div[2]"));

                                                            //get all Description
                                                            //string[] AllDescrpt = GetArrayElements("copy-small application-footnote");

                                                            //write Description
                                                            MySheet.Cells[row, 7] = elementDscrpt.Text;
                                                            tempcountforXpath++;
                                                        }
                                                        else
                                                        {

                                                        }
                                                        row++;
                                                        tempcount++;
                                                    }

                                                    //value is Group size
                                                    else
                                                    {
                                                        //write year
                                                        MySheet.Cells[row, 1] = linkTextYear;
                                                        //write make
                                                        MySheet.Cells[row, 2] = linkTextMake;
                                                        //write model
                                                        MySheet.Cells[row, 3] = linkTextModel;
                                                        //write engine
                                                        MySheet.Cells[row, 4] = linkTextEngine;
                                                        //write Group size
                                                        MySheet.Cells[row, 5] = value;
                                                        tempcount++;
                                                    }
                                                }
                                                MyBook.Save();

                                            }
                                            else
                                            {

                                            }

                                            //back previous page
                                            PropertiesCollection.driver.Navigate().Back();
                                        }
                                        System.Threading.Thread.Sleep(1000);

                                        IWebElement elementYears2 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                                        elementYears2.AsDropDown().SelectByText(linkTextYear);

                                        System.Threading.Thread.Sleep(1000);
                                        //select Make on DDL
                                        IWebElement elementMakes2 = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                                        elementMakes2.AsDropDown().SelectByText(linkTextMake);

                                        System.Threading.Thread.Sleep(1000);

                                        //select Model on DDL
                                        IWebElement elementModel2 = PropertiesCollection.driver.FindElement(By.Id("product_finder_model"));
                                        elementModel2.AsDropDown().SelectByText(linkTextModel);

                                        System.Threading.Thread.Sleep(1000);
                                    }
                                }
                                //System.Threading.Thread.Sleep(500);

                                //IWebElement elementYears3 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                                //elementYears3.AsDropDown().SelectByText(year);

                                //System.Threading.Thread.Sleep(500);
                                ////select Make on DDL
                                //IWebElement elementMakes3 = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                                //elementMakes3.AsDropDown().SelectByText(linkTextMake);

                                //System.Threading.Thread.Sleep(500);
                            }
                        }
                        //System.Threading.Thread.Sleep(500);

                        //IWebElement elementYears4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                        //elementYears4.AsDropDown().SelectByText(linkTextYear);

                        //System.Threading.Thread.Sleep(500);
                    }
                }
            }

            MyBook.SaveAs(foldername);
            MyBook.Close();
            MyApp.Quit();
        }

        public static void GetDataForYear()
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

            MySheet.Cells[1, 5] = "Group Size";
            MySheet.Cells[1, 6] = "Minimum Cold Cranking Amps";
            MySheet.Cells[1, 7] = "Description ";

            //row of data
            int row = 2;

            string path = Environment.CurrentDirectory;
            string foldername = Path.Combine(path, "output.xlsx");

            MyBook.SaveAs(foldername);
            //waitOnPage(5)
            //waitForPageUntilElementIsVisible(By.LinkText(linkTextRangeYear), 5);
            System.Threading.Thread.Sleep(1000);

            //select year on DDL
            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
            elementYears.AsDropDown().SelectByText(year);

            System.Threading.Thread.Sleep(1000);

            IWebElement elementMakes = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
            var listMakes = elementMakes.AsDropDown().Options;

            //get all Makes
            List<string> AllMakes = CommonMethods.ListStringfromIList(listMakes);

            foreach (string linkTextMake in AllMakes)
            {
                //skip Please choose Make
                if (String.IsNullOrEmpty(linkTextMake) || String.IsNullOrWhiteSpace(linkTextMake) || linkTextMake.Contains("Please Choose Make"))
                {
                    continue;
                }
                else
                {
                    System.Threading.Thread.Sleep(1000);
                    IWebElement elementMakes4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                    //select Make on DDL
                    elementMakes4.AsDropDown().SelectByText(linkTextMake);

                    System.Threading.Thread.Sleep(1000);

                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("product_finder_model"));
                    var listModels = elementModel.AsDropDown().Options;

                    //get all Models
                    List<string> AllModels = CommonMethods.ListStringfromIList(listModels);

                    foreach (string linkTextModel in AllModels)
                    {
                        //skip Please choose Model
                        if (String.IsNullOrEmpty(linkTextModel) || String.IsNullOrWhiteSpace(linkTextModel) || linkTextModel.Contains("Please Choose Model"))
                        {
                            continue;
                        }
                        else
                        {
                            System.Threading.Thread.Sleep(1000);
                            IWebElement elementModel4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_model"));
                            //select Model on DDL
                            elementModel4.AsDropDown().SelectByText(linkTextModel);

                            System.Threading.Thread.Sleep(1000);

                            IWebElement elementEngine = PropertiesCollection.driver.FindElement(By.Id("product_finder_engine"));
                            var listEngine = elementEngine.AsDropDown().Options;

                            //get all Models
                            List<string> AllEngine = CommonMethods.ListStringfromIList(listEngine);

                            foreach (string linkTextEngine in AllEngine)
                            {
                                //skip Please choose Engine
                                if (String.IsNullOrEmpty(linkTextEngine) || String.IsNullOrWhiteSpace(linkTextEngine) || linkTextEngine.Contains("Please Choose Engine"))
                                {
                                    continue;
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(1000);
                                    //select Model on DDL
                                    IWebElement elementEngine4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_engine"));
                                    elementEngine4.AsDropDown().SelectByText(linkTextEngine);
                                    System.Threading.Thread.Sleep(1000);

                                    //click Find button 
                                    IWebElement FindButton = PropertiesCollection.driver.FindElement(By.Id("find_button"));
                                    FindButton.Click();

                                    System.Threading.Thread.Sleep(1000);

                                    //check Element exists
                                    //do if exist
                                    if (CommonMethods.IsElementPresent(By.ClassName("label")))
                                    {

                                        int tempcount = 0;
                                        int tempcountforXpath = 1;
                                        //get all value
                                        List<string> AllValues = CommonMethods.GetListfromtheSameElements(By.ClassName("value"));
                                        foreach (string value in AllValues)
                                        {
                                            //value is Amps
                                            if (tempcount % 2 != 0)
                                            {
                                                //write Amps
                                                MySheet.Cells[row, 6] = value;

                                                //do if exist
                                                if (CommonMethods.IsElementPresent(By.XPath("/html/body/div/div/div/div[1]/section/div[3]/div[1]/div[2]")))
                                                {
                                                    
                                                    IWebElement elementDscrpt = PropertiesCollection.driver.FindElement(By.XPath("/html/body/div/div/div/div[1]/section/div[3]/div[" + tempcountforXpath + "]/div[2]"));

                                                    //get all Description
                                                    //string[] AllDescrpt = GetArrayElements("copy-small application-footnote");

                                                    //write Description
                                                    MySheet.Cells[row, 7] = elementDscrpt.Text;
                                                    tempcountforXpath++;
                                                }
                                                else
                                                {

                                                }
                                                row++;
                                                tempcount++;
                                            }

                                            //value is Group size
                                            else
                                            {
                                                //write year
                                                MySheet.Cells[row, 1] = year;
                                                //write make
                                                MySheet.Cells[row, 2] = linkTextMake;
                                                //write model
                                                MySheet.Cells[row, 3] = linkTextModel;
                                                //write engine
                                                MySheet.Cells[row, 4] = linkTextEngine;
                                                //write Group size
                                                MySheet.Cells[row, 5] = value;
                                                tempcount++;
                                            }
                                        }
                                        MyBook.Save();
                                        
                                    }
                                    else
                                    {

                                    }

                                    //back previous page
                                    PropertiesCollection.driver.Navigate().Back();
                                }
                                System.Threading.Thread.Sleep(1000);

                                IWebElement elementYears2 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                                elementYears2.AsDropDown().SelectByText(year);

                                System.Threading.Thread.Sleep(1000);
                                //select Make on DDL
                                IWebElement elementMakes2 = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                                elementMakes2.AsDropDown().SelectByText(linkTextMake);

                                System.Threading.Thread.Sleep(1000);

                                //select Model on DDL
                                IWebElement elementModel2 = PropertiesCollection.driver.FindElement(By.Id("product_finder_model"));
                                elementModel2.AsDropDown().SelectByText(linkTextModel);

                                System.Threading.Thread.Sleep(1000);
                            }
                        }
                        System.Threading.Thread.Sleep(1000);

                        IWebElement elementYears3 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                        elementYears3.AsDropDown().SelectByText(year);

                        System.Threading.Thread.Sleep(1000);
                        //select Make on DDL
                        IWebElement elementMakes3 = PropertiesCollection.driver.FindElement(By.Id("product_finder_make"));
                        elementMakes3.AsDropDown().SelectByText(linkTextMake);

                        System.Threading.Thread.Sleep(1000);
                    }
                }
                System.Threading.Thread.Sleep(1000);

                IWebElement elementYears4 = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                elementYears4.AsDropDown().SelectByText(year);

                System.Threading.Thread.Sleep(1000);
            }

            MyBook.SaveAs(foldername);
            MyBook.Close();
            MyApp.Quit();
        }
    }
}
