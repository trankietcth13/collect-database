using GetDataOnWeb_v01;
using OfficeOpenXml;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataOnWeb_v01
{
    public class Autocodes
    {
        public static List<string> listMake = new List<string>{
                "ACURA",
                "AUDI",
                "BMW",
                "BUICK",
                "CADILLAC",
                "CHEVROLET",
                "CHRYSLER",
                "DODGE",
                "EAGLE",
                "FORD",
                "GEO",
                "GMC",
                "HONDA",
                "HUMMER",
                "HYUNDAI",
                "INFINITY",
                "ISUZU",
                "JAGUAR",
                "KIA",
                "LAND ROVER",
                "LEXUS",
                "LINCOLN",
                "MAZDA",
                "MERCEDES-BENZ",
                "MERCURY",
                "MINI",
                "MITSUBISHI",
                "NISSAN",
                "OLDSMOBILE",
                "PONTIAC",
                "SAAB",
                "SATURN",
                "SCION",
                "SUBARU",
                "SUZUKI",
                "TOYOTA",
                "VW",
                "VOLVO"

        };

        //make from combobox
        public static string make { get; set; }

        //write header and first page to excel
        public static int WriteHeaderandFirstPage(string foldername, ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {

            //create header of YMME sheet
            MySheet.Cells[1, 1].Value = "Make";
            MySheet.Cells[1, 2].Value = "Code";
            MySheet.Cells[1, 3].Value = "Code Title";
            MySheet.Cells[1, 4].Value = "Code Description";
            MySheet.Cells[1, 5].Value = "Source Image";
            MySheet.Cells[1, 6].Value = "Importance Level";
            MySheet.Cells[1, 7].Value = "Difficulty Level";

            package.SaveAs(new System.IO.FileInfo(foldername));

            return rowYMME = WriteDataforMulPage(MySheet, rowYMME, colYMME, package);

        }

        //wirte data for multiple page
        public static int WriteDataforMulPage(ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            IWebElement all = PropertiesCollection.driver.FindElement(By.Id("scroller"));
            IList<IWebElement> allCod = all.FindElements(By.TagName("a"));
            List<string> allCodes = CommonMethods.ListStringfromIList(allCod);
            foreach (string code in allCodes)
            {
                System.Threading.Thread.Sleep(1000);

                IWebElement elementcode = PropertiesCollection.driver.FindElement(By.LinkText(code));
                elementcode.Click();
                System.Threading.Thread.Sleep(1000);

                //list data to write to excel
                List<string> listData = CommonMethods.GetListfromtheSameElements(By.ClassName("info_code"));
                //list tittles of code
                List<string> listTitleCode = CommonMethods.GetListfromtheSameElements(By.ClassName("code"));
                //value of Repair Difficulty Level
                string diffLevel = PropertiesCollection.driver.FindElement(By.XPath("/html/body/div/div[2]/div[4]/div[1]/span[4]")).Text;
                //value of Repair Importance Level
                string ImportanceLevel = PropertiesCollection.driver.FindElement(By.XPath("/html/body/div/div[2]/div[4]/div[1]/span[2]")).Text;

                //write name of make
                MySheet.Cells[rowYMME, 1].Value = make.ToUpper();
                //write name of code
                MySheet.Cells[rowYMME, 2].Value = code;
                //write Code Title
                MySheet.Cells[rowYMME, 3].Value = listTitleCode[0];

                //write link text of img
                IWebElement img = PropertiesCollection.driver.FindElement(By.ClassName("img_resize"));
                string srcimg = img.GetAttribute("src");
                MySheet.Cells[rowYMME, 5].Value = srcimg;

                //write value of Repair Importance Level
                MySheet.Cells[rowYMME, 6].Value = ImportanceLevel;
                //write value of Repair Difficulty Level
                MySheet.Cells[rowYMME, 7].Value = diffLevel;

                //wirte code Description and Data
                for (int i = 0; i < listTitleCode.Count(); i++)
                {
                    if (i == 0 || listTitleCode[i].Contains("Need more help?") || listTitleCode[i].Contains("Help AutoCodes.com") || listTitleCode[i].Contains("Related Information") || listTitleCode[i].Contains("Comments"))
                    {
                        continue;
                    }
                    if (listTitleCode[i].Contains("Description"))
                    {
                        //write code Description
                        MySheet.Cells[rowYMME, 4].Value = listData[i-1];
                    }
                    else
                    {
                        //column in excel = null
                        if (MySheet.Cells[1, colYMME].Value == null)
                        {
                            //write tittle
                            MySheet.Cells[1, colYMME].Value = listTitleCode[i];
                            //write  Data
                            MySheet.Cells[rowYMME, colYMME].Value = listData[i - 1];
                        }
                        else
                        {
                            //cloumn in excel = tittle
                            if (MySheet.Cells[1, colYMME].Value.ToString().ToLower().Contains(listTitleCode[i].ToLower()))
                            {
                                //write  Data
                                MySheet.Cells[rowYMME, colYMME].Value = listData[i - 1];
                            }
                            else
                            {
                                for (int j = 8; j < 100; j++)
                                {
                                    if (MySheet.Cells[1, j].Value == null || MySheet.Cells[1, j].Value.ToString().ToLower().Contains(listTitleCode[i].ToLower()))
                                    {
                                        colYMME = j;
                                        break;
                                    }
                                }
                                //write tittle
                                MySheet.Cells[1, colYMME].Value = listTitleCode[i];
                                //write  Data
                                MySheet.Cells[rowYMME, colYMME].Value = listData[i - 1];
                                colYMME = colYMME - 1;
                            }
                        }
                    }
                    colYMME++;
                }
                package.Save();
                colYMME = 8;
                rowYMME++;
                PropertiesCollection.driver.Navigate().Back();
            }
            return rowYMME;
        }

        //get data from Make which is choose from combobox
        public static void GetDatabyMake(string href)
        {
            string[] temp = href.Split('/');
            string make = temp[4];

            string numPage = temp[temp.Count() - 1];
            if (make.Equals(numPage))
            {
                numPage = "all";
            }
            string path = Environment.CurrentDirectory;
            string namefile = "outputAutoCodes_" + make + "_" + numPage + ".xlsx";
            string foldername = Path.Combine(path, namefile);

            System.Threading.Thread.Sleep(500);

            int rowYMME = 2;
            int colYMME = 8;
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                workbook.Worksheets.Add("YMME");
                var MySheet = workbook.Worksheets[1];
                workbook.Worksheets.Add("Database");

                //write data to excel for first page of Make
                rowYMME = WriteHeaderandFirstPage(foldername, MySheet, rowYMME, colYMME, package);



                //element of switch page
                IWebElement all = PropertiesCollection.driver.FindElement(By.Id("pag"));
                string numtmp = all.Text;
                string[] numtringtmp = numtmp.Split('f');
                numtmp = numtringtmp[1].Trim();
                string[] num2 = numtmp.Split(' ');
                string numofPage = num2[0].Trim();
                int numofpage = Convert.ToInt32(numofPage);
                int numcountPage; 
                if(numPage.Equals("all"))
                {
                    numcountPage = 1;
                }
                else
                {
                    numcountPage = Convert.ToInt32(numPage);
                }
                for(; numcountPage < numofpage; numcountPage++)
                {
                    //IList elements Next/Prev/First/Last Page
                    //IList<IWebElement> allPages = all.FindElements(By.TagName("a"));
                    System.Threading.Thread.Sleep(1000);
                    //click Next Page
                    IWebElement elementNextPage = PropertiesCollection.driver.FindElement(By.LinkText("Next >"));
                    elementNextPage.Click();
                    System.Threading.Thread.Sleep(1000);

                    rowYMME = WriteDataforMulPage(MySheet, rowYMME, colYMME, package);
                }
                
            }

        }

        //get data for each page from SubLink web input
        public static void GetDatabyPageCode(string href)
        {

            string[] temp = href.Split('/');
            string make = temp[4];

            string numPage = temp[temp.Count() - 1];
            if (make.Equals(numPage))
            {
                numPage = "01";
            }

            string path = Environment.CurrentDirectory;
            string namefile = "outputAutoCodes_" + make + "_" + numPage + ".xlsx";
            string foldername = Path.Combine(path, namefile);

            System.Threading.Thread.Sleep(500);

            int rowYMME = 2;
            int colYMME = 8;
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                workbook.Worksheets.Add("YMME");
                var MySheet = workbook.Worksheets[1];
                workbook.Worksheets.Add("Database");

                //write data to excel for first page of Make
                rowYMME = WriteHeaderandFirstPage(foldername, MySheet, rowYMME, colYMME, package);
            }
        }
    }
}
