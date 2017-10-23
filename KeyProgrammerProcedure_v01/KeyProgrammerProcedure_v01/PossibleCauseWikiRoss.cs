using OfficeOpenXml;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataOnWeb_v01
{
    public class PossibleCauseWikiRoss
    {
        public static void GetPosCauseForPage()
        {

            string path = Environment.CurrentDirectory;
            string namefile = "outputPossibleCause.xlsx";
            string foldername = Path.Combine(path, namefile);

            System.Threading.Thread.Sleep(500);

            int rowYMME = 2;
            int colYMME = 8;
            using (var package = new ExcelPackage())
            {
                var workbook = package.Workbook;
                workbook.Worksheets.Add("ListPossible");
                var MySheet = workbook.Worksheets[1];
                workbook.Worksheets.Add("Database");
                var MySheet2 = workbook.Worksheets[2];

                //write data to excel for first page of Make
                WriteHeader(foldername, MySheet, rowYMME, colYMME, package);
                rowYMME = WriteData(MySheet, rowYMME, colYMME, package);
            }
        }
        public static void WriteHeader(string foldername, ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            //create header of YMME sheet
            MySheet.Cells[1, 1].Value = "Code";
            MySheet.Cells[1, 2].Value = "Title01";
            MySheet.Cells[1, 3].Value = "Title02";
            MySheet.Cells[1, 4].Value = "Possible Symptoms";
            MySheet.Cells[1, 5].Value = "Possible Causes";
            MySheet.Cells[1, 6].Value = "Possible Solutions";
            MySheet.Cells[1, 7].Value = "Special Notes";

            package.SaveAs(new System.IO.FileInfo(foldername));

        }

        public static int WriteData(ExcelWorksheet MySheet, int rowYMME, int colYMME, ExcelPackage package)
        {
            //element include all link
            IWebElement all = PropertiesCollection.driver.FindElement(By.ClassName("mw-content-ltr"));
            //list elements all link
            IList<IWebElement> allCod = all.FindElements(By.TagName("a"));
            //list string of all link
            List<string> allCodes = CommonMethods.ListStringfromIList(allCod);
            foreach (string code in allCodes)
            {
                if (code.Contains("next"))
                {
                    continue;
                }
                System.Threading.Thread.Sleep(1000);
                IWebElement elementcode = PropertiesCollection.driver.FindElement(By.LinkText(code));
                elementcode.Click();
                System.Threading.Thread.Sleep(1000);

                //element include data and title
                IWebElement ElementofDataTitle = PropertiesCollection.driver.FindElement(By.Id("mw-content-text"));
                //list elements of title01
                IList<IWebElement> listTittle01 = ElementofDataTitle.FindElements(By.TagName("h2"));
                //list elements of title02
                IList<IWebElement> listTittle02 = ElementofDataTitle.FindElements(By.TagName("h3"));
                //list elements of header
                IList<IWebElement> listHeader = ElementofDataTitle.FindElements(By.TagName("h4"));
                //list elements of data
                IList<IWebElement> listData = ElementofDataTitle.FindElements(By.TagName("ul"));

                int title01 = 0;
                int title02 = 0;
                int header = 0;
                foreach (var data in listData)
                {
                    if (listTittle01.Count != 0)
                    {
                        //write title 1
                        MySheet.Cells[rowYMME, 2].Value = listTittle01[title01].Text;
                        title01++;
                    }
                    if (listTittle02.Count != 0)
                    {
                        //write title 1
                        MySheet.Cells[rowYMME, 3].Value = listTittle02[title02].Text;
                        title02++;
                    }
                    switch(listHeader[header].Text)
                    {
                        case "Possible Symptoms":
                            MySheet.Cells[rowYMME, 4].Value = data.Text;
                            break;
                        default:
                            break;
                    }

                }

                PropertiesCollection.driver.Navigate().Back();
            }
            return rowYMME;
        }
    }
}
