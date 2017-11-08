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
using OpenQA.Selenium.Support.UI;

namespace KeyProgrammerProcedure_v01
{
    //public static class SeleniumExtensions
    //{
    //    public static void WaitForNavigation(this IWebDriver driver)
    //    {
    //        var wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30));
    //        wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
    //    }
    //}
    class GetDataSourcebci
    {

        

        public static List<string> listMake = new List<string>
        {
            //"ACURA", //done
           // "LINCOLN",//done
           // "MAZDA", //done
           // "MERCURY",
           // "ACURA", //done
           // "AUDI", //done
          //  "BMW", //done
           // "BUICK", //done
           // "CADILLAC", //done
          //  "CHEVROLET", //done
          //  "CHRYSLER", //done
          //  "DAEWOO", //done
          //  "FORD", //done
          //  "DODGE", //done
          //  "EAGLE",
          //  "GEO",
           // "GMC",
            //"HONDA", //done
           // "HUMMER",
           // "HYUNDAI",//done
           // "INFINITI", //done
           // "ISUZU",
           // "JAGUAR",//done
           // "JEEP",//done
           // "KIA", //done
           // "LAND ROVER",//done
          //  "LEXUS",//done
          //  "MERCEDES",//done
          //   "MINI", //done
          // "MITSUBISHI",
          // "NISSAN",//done
          // // "OLDSMOBILE",
          // // "PLYMOUTH",
          ////  "PONTIAC",
          // // "PORSCHE",
          // // "SAAB",
          // // "SATURN",
           //"SCION",
          // // "SMART",
          // // "SUBARU",
          // // "SUZUKI",
           "TOYOTA"
          //  "VOLKSWAGEN",
          //  "VOLVO"
        };

        public static void WriteDataExcel()
        {
            

               try
               {
                   string path = Environment.CurrentDirectory;
                   string namefile = "SourceBCIData_.xlsx";
                   string foldername = Path.Combine(path, namefile);
                   int rowYMME = 2;
                   int colYMME = 8;
                   using (var package = new ExcelPackage())
                   {
                       var workbook = package.Workbook;
                       workbook.Worksheets.Add("Database");
                       var MySheet = workbook.Worksheets[1];

                       //write data to excel for first page of Make
                       WriteHeader(foldername, MySheet, rowYMME, colYMME, package);
                       rowYMME = GetdataBCI(MySheet, rowYMME, colYMME, package);
                       System.Threading.Thread.Sleep(1000);
                }
                   
               }

               catch (Exception ex)
               {
                   Console.WriteLine("The process failed: {0}", ex.ToString());
                  
               }
           
        }

        public static void WriteHeader(string foldername, ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {

            try
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
            catch (Exception ex)
            {
                Console.WriteLine("The process failed: {0}", ex.ToString());
            }                           
        }

        public static int GetdataBCI(ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            bool staleElement = true;
            
                try
                {
                    #region collect database CCA
                    //select Makes on DDL
                    WebDriverWait wait = new WebDriverWait(PropertiesCollection.driver, TimeSpan.FromSeconds(5));
                    IWebElement elementMakes = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
                    var listMakes = elementMakes.AsDropDown().Options;
                    //get all Makes
                    List<string> AllMakes = CommonMethods.ListStringfromIList(listMakes);
                    int row = 2;
                    foreach (string _listMakes in AllMakes)
                    {
                        //skip Please choose Make
                        if (String.IsNullOrEmpty(_listMakes) || String.IsNullOrWhiteSpace(_listMakes) || _listMakes.Contains("-select-")) //|| _listMakes.Contains("Acura") || _listMakes.Contains("Audi") || _listMakes.Contains("BMW") || _listMakes.Contains("Hummer") || _listMakes.Contains("Cadillac") || _listMakes.Contains("Buick") || _listMakes.Contains("Chevrolet") || _listMakes.Contains("Chrysler") || _listMakes.Contains("Dodge") || _listMakes.Contains("Daewoo") || _listMakes.Contains("Eagle") || _listMakes.Contains("Geo") || _listMakes.Contains("GMC") || _listMakes.Contains("Honda") || _listMakes.Contains("Hyundai") || _listMakes.Contains("Infiniti") || _listMakes.Contains("Isuzu") || _listMakes.Contains("Jaguar") || _listMakes.Contains("Jeep") || _listMakes.Contains("Ford") || _listMakes.Contains("Kia") || _listMakes.Contains("Lexus") || _listMakes.Contains("Land Rover"))
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

                                    IWebElement elementMakes4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1")); //PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
                                 //select Make on DDL
                                    elementMakes4.AsDropDown().SelectByText(_listMakes);
                                    
                                    //PropertiesCollection.driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(4));
                                    System.Threading.Thread.Sleep(16000);
                                    IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1")); //PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                                    var listYears = elementYears.AsDropDown().Options;

                                    //get all Models
                                    List<string> AllYears = CommonMethods.ListStringfromIList(listYears);

                                    foreach (string _listYears in AllYears)
                                    {
                                        //skip Please choose Make
                                        if (String.IsNullOrEmpty(_listYears) || String.IsNullOrWhiteSpace(_listYears) || _listYears.Contains("-select-"))
                                        {
                                            continue;
                                        }
                                        int year = Convert.ToInt32(_listYears);
                                        if (year < 1996 || year > 2006)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            //System.Threading.Thread.Sleep(3000);
                                            IWebElement elemenYears4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                                            //select Make on DDL
                                            elemenYears4.AsDropDown().SelectByText(_listYears);
                                            
                                            System.Threading.Thread.Sleep(16000);

                                            IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1")); //PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                            var listModels = elementModel.AsDropDown().Options;

                                            //get all Models
                                            System.Threading.Thread.Sleep(1000);
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
                                                    IWebElement elementModel4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                                    //select Model on DDL
                                                    elementModel4.AsDropDown().SelectByText(_listModels);

                                                    System.Threading.Thread.Sleep(16000);

                                                    #region Get BCI & CCA                                           
                                                    //List data
                                                    IWebElement FindTagTd = PropertiesCollection.driver.FindElement(By.Id("MainContent_gvApplication"));
                                                    IList<IWebElement> TagTd_table = FindTagTd.FindElements(By.TagName("td"));
                                                    List<string> AlltagTd = CommonMethods.ListStringfromIList(TagTd_table);

                                                    //Find and click button plus
                                                    IList<IWebElement> TagTd_img = FindTagTd.FindElements(By.TagName("img"));
                                                    List<string> allTagimg = CommonMethods.ListStringfromIList(TagTd_img);
                                                    int img_num = 0;
                                                    int count2 = 0;
                                                    foreach (var itemTD in allTagimg)
                                                    {
                                                        count2++;
                                                        if (count2 > allTagimg.Count)
                                                        {
                                                            break;
                                                        }
                                                        if (itemTD.Contains("hide-button"))
                                                        {
                                                            continue;
                                                        }
                                                        else
                                                        {
                                                            IWebElement findBtn = PropertiesCollection.driver.FindElement(By.Id("MainContent_gvApplication_btnDetail_" + img_num));
                                                            string imgSrc = findBtn.GetAttribute("src");
                                                            if (imgSrc.Contains("expand-button"))
                                                            {
                                                                findBtn.Click();
                                                                System.Threading.Thread.Sleep(1000);
                                                                img_num++;
                                                            }
                                                        }
                                                    }
                                                    // -----End-----Find and click plus button ------

                                                    //List database 
                                                    int colofTagTd = 4;
                                                    int count = 0;

                                                    IWebElement FindTagTd2 = PropertiesCollection.driver.FindElement(By.ClassName("gridview"));
                                                    IList<IWebElement> TagTd_table2 = FindTagTd2.FindElements(By.TagName("td"));
                                                    List<string> AlltagTd2 = CommonMethods.ListStringfromIList(TagTd_table2);

                                                    #region Collect data
                                                    foreach (var item_TagTd in AlltagTd2)
                                                    {
                                                        count++;
                                                        if (count >= AlltagTd2.Count)
                                                        {
                                                            break;
                                                        }
                                                        if (String.IsNullOrEmpty(item_TagTd) || item_TagTd.ToLower().Equals(_listMakes.ToLower()) || item_TagTd.ToLower().Equals(_listModels.ToLower()) || item_TagTd.ToLower().ToString().Contains(_listMakes) || item_TagTd.ToLower().ToString().Contains(_listModels))
                                                        {
                                                            continue;
                                                        }
                                                        else
                                                        {
                                                            if (item_TagTd.Trim().ToLower().Equals(_listYears))
                                                            {
                                                                //write make
                                                                MySheet.Cells[row, 1].Value = _listMakes;
                                                                //write year
                                                                MySheet.Cells[row, 2].Value = _listYears;
                                                                //write model
                                                                MySheet.Cells[row, 3].Value = _listModels;
                                                                //write engine
                                                                MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count];
                                                                package.Save();
                                                                colofTagTd++;
                                                                continue;
                                                            }
                                                            else
                                                            {
                                                                if (item_TagTd.Trim().ToLower().Contains("notes"))
                                                                {
                                                                    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 1]; //get BCI
                                                                    colofTagTd++;
                                                                    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 2]; //get CCA 
                                                                                                                                 //colofTagTd = 4;
                                                                    package.Save();

                                                                    for (int i = 1; i < AlltagTd2.Count; i++)
                                                                    {
                                                                        //size of all  tagTd > index of td after td contains Notes
                                                                        if (AlltagTd2.Count > (AlltagTd2.IndexOf(item_TagTd) + 1 + (5 * i)))
                                                                        {
                                                                            //text of td after td Notes == null
                                                                            if (AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 1 + (5 * i)] == null)
                                                                            {
                                                                                //text of td after td Notes = make
                                                                                if (AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 2 + (5 * i)].ToLower().Equals(_listMakes.ToLower()))
                                                                                {
                                                                                    break;
                                                                                }
                                                                            }

                                                                            if (AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 2 + (5 * i)].ToLower().Equals(_listMakes.ToLower()))
                                                                            {
                                                                                break;
                                                                            }

                                                                            colofTagTd++;
                                                                            MySheet.Cells[row, colofTagTd].Value = AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 2 + (5 * i)]; //get BCI(n)
                                                                            colofTagTd++;
                                                                            MySheet.Cells[row, colofTagTd].Value = AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 3 + (5 * i)]; //get CCA(n)
                                                                            package.Save();
                                                                            continue;
                                                                        }
                                                                        else
                                                                        {
                                                                            break;
                                                                        }
                                                                    }

                                                                    row++;
                                                                    colofTagTd = 4;
                                                                    continue;
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion
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
                   
                    #endregion
                    staleElement = false;
                }
                catch (StaleElementReferenceException ex)
                {
                    staleElement = true;
                }
            
            return rowYMME;

        }
    }
}