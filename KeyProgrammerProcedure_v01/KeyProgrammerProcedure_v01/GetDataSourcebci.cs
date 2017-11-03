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
            "Ford",
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
            foreach (string _listMakes in AllMakes)
            {
                //skip Please choose Make
                if (String.IsNullOrEmpty(_listMakes) || String.IsNullOrWhiteSpace(_listMakes) || _listMakes.Contains("-select-") || _listMakes.Contains("Acura") || _listMakes.Contains("Audi" +
                    "") || _listMakes.Contains("BMW") || _listMakes.Contains("Hummer") || _listMakes.Contains("Cadillac") || _listMakes.Contains("Buick") || _listMakes.Contains("Chevrolet") || _listMakes.Contains("Chrysler") || _listMakes.Contains("Dodge") || _listMakes.Contains("Eagle") || _listMakes.Contains("Geo") || _listMakes.Contains("GMC") || _listMakes.Contains("Honda") || _listMakes.Contains("Hyundai") || _listMakes.Contains("Infiniti") || _listMakes.Contains("Isuzu") || _listMakes.Contains("Jaguar"))//Chrysler|| _listMakes.Contains("Acura") 
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
                            //MySheet.Cells[row, 1].Value = _listMakes;
                            //package.Save();
                            System.Threading.Thread.Sleep(4000);

                            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                            var listYears = elementYears.AsDropDown().Options;

                            //get all Models
                            List<string> AllYears = CommonMethods.ListStringfromIList(listYears);

                            foreach (string _listYears in AllYears)
                            {
                                //skip Please choose Make
                                if (String.IsNullOrEmpty(_listYears) || String.IsNullOrWhiteSpace(_listYears) || _listYears.Contains("-select-"))//|| _listYears.Contains("2015") || _listYears.Contains("2016")|| _listYears.Contains("2013") || _listYears.Contains("2014") || _listYears.Contains("2012") || _listYears.Contains("2011") || _listYears.Contains("2010")) // || _listYears.Contains("2009") || _listYears.Contains("2008")) // || _listYears.Contains("2009")|| _listYears.Contains("2015") || _listYears.Contains("2014") || _listYears.Contains("2013") || _listYears.Contains("2012") || _listYears.Contains("2011") || _listYears.Contains("2010") || _listYears.Contains("2009")) //|| _listYears.Contains("2013")) //|| _listYears.Contains("2016")
                                {
                                    continue;
                                }
                                int year = Convert.ToInt32(_listYears);
                                if (year < 1986 || year > 2016)
                                {
                                    continue;
                                }
                                else
                                {
                                    System.Threading.Thread.Sleep(3000);
                                    IWebElement elemenYears4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                                    //select Make on DDL
                                    elemenYears4.AsDropDown().SelectByText(_listYears);

                                    //MySheet.Cells[row, 2].Value = _listYears;
                                    //package.Save();

                                    System.Threading.Thread.Sleep(5000);

                                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                    var listModels = elementModel.AsDropDown().Options;

                                    //get all Models
                                    List<string> AllModels = CommonMethods.ListStringfromIList(listModels);
                                    foreach (string _listModels in AllModels)
                                    {
                                        //skip Please choose Model
                                        if (String.IsNullOrEmpty(_listModels) || String.IsNullOrWhiteSpace(_listModels) || _listModels.Contains("-select-") || _listModels.Contains("Compass"))//
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            System.Threading.Thread.Sleep(1000);
                                            IWebElement elementModel4 = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                                            //select Model on DDL
                                            elementModel4.AsDropDown().SelectByText(_listModels);
                                            //MySheet.Cells[row, 3].Value = _listModels;
                                            //package.Save();
                                            System.Threading.Thread.Sleep(5000);

                                    #region Get BCI & CCA                                           
                                            //List data
                                            IWebElement FindTagTd = PropertiesCollection.driver.FindElement(By.ClassName("gridview"));
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
                                            IList<IWebElement> TagTd_table2 = FindTagTd.FindElements(By.TagName("td"));
                                            List<string> AlltagTd2 = CommonMethods.ListStringfromIList(TagTd_table2);
                                           
                                            #region Collect data
                                            foreach (var item_TagTd in AlltagTd2)
                                            {
                                                count++;
                                                if (count >= AlltagTd2.Count)
                                                {
                                                    break;
                                                }
                                                if (String.IsNullOrEmpty(item_TagTd) || item_TagTd.ToLower().Equals(_listMakes.ToLower()) || item_TagTd.ToLower().Equals(_listModels.ToLower()) || item_TagTd.ToLower().ToString().Contains(_listMakes) || item_TagTd.ToLower().ToString().Contains(_listModels)) //tempEngine == 1||
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
                                                            #region Note
                                                            //if(AlltagTd2.Count <= 6)
                                                            //{
                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 6]; //get BCI
                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 7]; //get CCA 
                                                            //    continue;
                                                            //}

                                                            //if(AlltagTd2.Count <= 22 )
                                                            //{
                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 12]; //get BCI
                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 13]; //get CCA
                                                            //    colofTagTd++;

                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 16]; //get BCI
                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 17]; //get CCA  

                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 21]; //get BCI
                                                            //    colofTagTd++;
                                                            //    MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 22]; //get CCA  

                                                            //}

                                                            //if(AlltagTd2.Count > 22|| AlltagTd2.Count <= 82)
                                                            //{
                                                            //    if( AlltagTd2[count + 26] != null|| AlltagTd2[count + 27] != null|| AlltagTd2[count + 31] != null|| AlltagTd2[count + 32] != null)
                                                            //    {


                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 26]; //get BCI
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 27]; //get CCA  

                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 31]; //get BCI
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 32]; //get CCA 
                                                            //    }
                                                            //    if(AlltagTd2[count + 36] != null|| AlltagTd2[count + 37] != null)
                                                            //    {
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 36]; //get BCI
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 37]; //get CCA 
                                                            //    }

                                                            //    if (AlltagTd2[count + 41] != null || AlltagTd2[count + 42] != null)
                                                            //    {
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 41]; //get BCI
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 42]; //get CCA 
                                                            //    }

                                                            //}
                                                            #endregion
                                                            #region Note 2
                                                            //IWebElement findBCI_CCA_01 = PropertiesCollection.driver.FindElement(By.ClassName("subgridview"));
                                                            //IList<IWebElement> findBCI_CCA_02 = findBCI_CCA_01.FindElements(By.TagName("td"));

                                                            //IList<IWebElement> findtag_a = findBCI_CCA_01.FindElements(By.TagName("a"));
                                                            //IList<string> list_tag_a = CommonMethods.ListStringfromIList(findtag_a);

                                                            //IList<string> listBCI_CCA = CommonMethods.ListStringfromIList(findBCI_CCA_02);
                                                            //int count3 = 0;

                                                            //foreach (var itemBCI_CCA in list_tag_a)
                                                            //{
                                                            //    count3++;
                                                            //    if (count3 > list_tag_a.Count)
                                                            //    {
                                                            //        break;
                                                            //    }
                                                            //    if (String.IsNullOrEmpty(itemBCI_CCA) || String.IsNullOrWhiteSpace(itemBCI_CCA)|| itemBCI_CCA == item_TagTd || itemBCI_CCA == AlltagTd2[count]||itemBCI_CCA.Contains("-"))
                                                            //    {
                                                            //        continue;
                                                            //    }
                                                            //    else
                                                            //    {
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 1]; //get BCI
                                                            //        colofTagTd++;
                                                            //        MySheet.Cells[row, colofTagTd].Value = AlltagTd2[count + 2]; //get CCA                                                   
                                                            //        colofTagTd = 4;
                                                            //        package.Save();
                                                            //        continue;

                                                            //    }
                                                            //}
                                                            #endregion

                                                            for(int i =1; i<=AlltagTd2.Count; i++)
                                                            {                                                                
                                                                if (String.IsNullOrEmpty((AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 5 * (i)])) ||(AlltagTd2[AlltagTd2.IndexOf(item_TagTd)+5*(i)]) != _listMakes)//|| (AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + 5 * (i)])!=null)
                                                                {                                                                        
                                                                        //if (AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + (5 * i) + 2].ToString() == _listMakes || AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + (5 * i) + 3].ToString() == _listMakes)// || AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + (5 * i) + 3].ToString() == _listModels|| AlltagTd2[AlltagTd2.IndexOf(item_TagTd) + (5 * i) + 3].ToString() == _listYears)
                                                                        //{
                                                                        //    break;
                                                                        //}
                                                                        //else
                                                                        //{
                                                                            colofTagTd++;
                                                                            MySheet.Cells[row, colofTagTd].Value = AlltagTd2[AlltagTd2.IndexOf(item_TagTd)  + 6];
                                                                            colofTagTd++;
                                                                            MySheet.Cells[row, colofTagTd].Value = AlltagTd2[AlltagTd2.IndexOf(item_TagTd)  + 7];
                                                                            package.Save();
                                                                            continue;
                                                                       // }
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
            return rowYMME;
        }
    }
}