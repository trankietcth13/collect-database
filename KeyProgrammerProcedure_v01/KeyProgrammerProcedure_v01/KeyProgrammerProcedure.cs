using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Net;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing;

namespace GetDataOnWeb_v01
{
    class KeyProgrammerProcedure
    {
        //public static void NavigateandHandle(string url)
        //{
        //    //Navigate to Web
        //    PropertiesCollection.driver.Navigate().GoToUrl(url);
        //    Console.WriteLine("Opened URL");

        //    //Handle();
        //}
        public static string make;
        public static void GetImgforMake()
        {
            //get makes

            string path = Environment.CurrentDirectory;
            string foldername = Path.Combine(path, make);
            Directory.CreateDirectory(foldername);


            //click make
            IWebElement elementMake = PropertiesCollection.driver.FindElement(By.LinkText(make));
            elementMake.Click();

            //get models
            List<string> AllModelLinks = GetArrayLinks();

            foreach (string linkTextModel in AllModelLinks)
            {
                if (String.IsNullOrEmpty(linkTextModel) || String.IsNullOrWhiteSpace(linkTextModel))
                {
                    continue;
                }
                else
                {
                  
                    string foldernameModel = Path.Combine(foldername, linkTextModel.Replace("/", "-").Replace(">", "-"));
                    Directory.CreateDirectory(foldernameModel);
                    //click model
                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.LinkText(linkTextModel));
                    elementModel.Click();

                    //get engines
                    List<string> AllEnginesLinks = GetArrayLinks();
                    #region Engine or Not
                    foreach (string linkTextEngine in AllEnginesLinks)
                    {
                        if (String.IsNullOrEmpty(linkTextEngine) || String.IsNullOrWhiteSpace(linkTextEngine))
                        {
                            continue;
                        }
                        else
                        {
                            string foldernameEngine = Path.Combine(foldernameModel, linkTextEngine.Replace("/", "-").Replace(">", "-"));
                            Directory.CreateDirectory(foldernameEngine);
                            //click engine
                            IWebElement elementEngine = PropertiesCollection.driver.FindElement(By.LinkText(linkTextEngine));
                            elementEngine.Click();

                            ////write make
                            //MySheet.Cells[i, 1] = linkTextMake;
                            ////write model
                            //MySheet.Cells[i, 2] = linkTextModel;

                            //if not have engine
                            if (linkTextEngine.Contains(">"))
                            {
                                //write engine
                                //MySheet.Cells[i, 4] = linkTextEngine;

                                //check Element exists
                                //CheckElementExist(i, MySheet);
                                //check img exists
                                int countImg = 1;
                                By byImg = By.ClassName("locatorimage");
                                var elementImgSource = PropertiesCollection.driver.FindElements(byImg).Count >= 1 ? PropertiesCollection.driver.FindElement(byImg) : null;
                                if (elementImgSource != null)
                                {

                                    IList<IWebElement> listImg = PropertiesCollection.driver.FindElements(By.ClassName("locatorimage"));
                                    foreach (IWebElement elementimgSource in listImg)
                                    {
                                        // MySheet.Cells[i, colImg] = "\n" + elementimgSource.GetAttribute("src");

                                        SaveImage(foldernameEngine, elementimgSource, make, linkTextModel, linkTextEngine, " ", countImg);
                                        countImg++;

                                    }

                                }

                            }
                            else
                            {
                                //get range year
                                List<string> AllRangeYearLinks = GetArrayLinks();

                                #region Range Year
                                foreach (string linkTextRangeYear in AllRangeYearLinks)
                                {
                                    if (String.IsNullOrEmpty(linkTextRangeYear) || String.IsNullOrWhiteSpace(linkTextRangeYear))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        string foldernameRangeYear = Path.Combine(foldernameEngine, linkTextRangeYear.Replace("/", "-").Replace(">", "-"));
                                        Directory.CreateDirectory(foldernameRangeYear);
                                        //click engine
                                        IWebElement elementRangeYear = PropertiesCollection.driver.FindElement(By.LinkText(linkTextRangeYear));

                                        elementRangeYear.Click();

                                        ////write engine
                                        //MySheet.Cells[i, 3] = linkTextEngine;
                                        ////write range year
                                        //MySheet.Cells[i, 4] = linkTextRangeYear;

                                        //check Element exists
                                        //CheckElementExist(i, MySheet);

                                        int colImg = 7;
                                        int countImg = 1;

                                        //check img exists
                                        By byImg = By.ClassName("locatorimage");
                                        var elementImgSource = PropertiesCollection.driver.FindElements(byImg).Count >= 1 ? PropertiesCollection.driver.FindElement(byImg) : null;
                                        if (elementImgSource != null)
                                        {

                                            IList<IWebElement> listImg = PropertiesCollection.driver.FindElements(By.ClassName("locatorimage"));
                                            foreach (IWebElement elementimgSource in listImg)
                                            {
                                                // MySheet.Cells[i, colImg] = "\n" + elementimgSource.GetAttribute("src");

                                                SaveImage(foldernameRangeYear, elementimgSource, make, linkTextModel, linkTextEngine, linkTextRangeYear, countImg);
                                                countImg++;
                                                colImg++;
                                            }

                                        }

                                        PropertiesCollection.driver.Navigate().Back();
                                    }

                                }
                                #endregion
                            }
                            PropertiesCollection.driver.Navigate().Back();
                        }

                    }
                    #endregion
                    PropertiesCollection.driver.Navigate().Back();
                }
            }
            //PropertiesCollection.driver.Navigate().Back();
        }
        public static void GetAllData()
        {
            List<string> AllMakeLinks = GetArrayLinks();

            //waitOnPage(5)
            //waitForPageUntilElementIsVisible(By.LinkText(linkTextRangeYear), 5);
            //System.Threading.Thread.Sleep(3000);

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
            MySheet.Cells[1, 1] = "MAKE";
            MySheet.Cells[1, 2] = "MODEL";
            MySheet.Cells[1, 3] = "ENGINE";
            MySheet.Cells[1, 4] = "RANGE YEAR";
            MySheet.Cells[1, 5] = "KEY PROGRAMMER";
            MySheet.Cells[1, 6] = "Tittle";
            MySheet.Cells[1, 7] = "IMG ";
            MySheet.Cells[1, 8] = "IMG 2";

            //list all make to build
            List<string> listMake = new List<string>{
                "ACURA",
                "AUDI",
                "BMW",
                "BUICK",
                "CADILLAC",
                "CHEVROLET",
                "FORD",
                "GEO",
                "GMC",
                "HONDA",
                "HUMMER",
                "HYUNDAI",
                "INFINITY",
                "KIA",
                "LEXUS",
                "LINCOLN",
                "MAZDA",
                "MERCEDES",
                "MERCURY",
                "MITSUBISHI",
                "NISSAN",
                "OLDSMOBILE",
                "PONTIAC",
                "SUZUKI",
                "TOYOTA",
                "VW"};

            int i = 2;
            //get makes
            foreach (string linkTextMake in AllMakeLinks)
            {
                if (String.IsNullOrEmpty(linkTextMake) || String.IsNullOrWhiteSpace(linkTextMake))
                {
                    continue;
                }
                else
                {
                    string path = Environment.CurrentDirectory;
                    string foldername = Path.Combine(path, linkTextMake);
                    Directory.CreateDirectory(foldername);

                    for (int k = 0; k < listMake.Count(); k++)
                    {
                        if (listMake[k].Contains(linkTextMake))
                        {
                            //click make
                            IWebElement elementMake = PropertiesCollection.driver.FindElement(By.LinkText(linkTextMake));
                            elementMake.Click();

                            //get models
                            List<string> AllModelLinks = GetArrayLinks();

                            foreach (string linkTextModel in AllModelLinks)
                            {
                                if (String.IsNullOrEmpty(linkTextModel) || String.IsNullOrWhiteSpace(linkTextModel))
                                {
                                    continue;
                                }
                                else
                                {
                                    string foldernameModel = Path.Combine(foldername, linkTextModel);
                                    Directory.CreateDirectory(foldernameModel);
                                    //click model
                                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.LinkText(linkTextModel));
                                    elementModel.Click();

                                    //get engines
                                    List<string> AllEnginesLinks = GetArrayLinks();
                                    #region Engine or Not
                                    foreach (string linkTextEngine in AllEnginesLinks)
                                    {
                                        if (String.IsNullOrEmpty(linkTextEngine) || String.IsNullOrWhiteSpace(linkTextEngine))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            string foldernameEngine = Path.Combine(foldernameModel, linkTextEngine);
                                            Directory.CreateDirectory(foldernameEngine);
                                            //click engine
                                            IWebElement elementEngine = PropertiesCollection.driver.FindElement(By.LinkText(linkTextEngine));
                                            elementEngine.Click();

                                            //write make
                                            MySheet.Cells[i, 1] = linkTextMake;
                                            //write model
                                            MySheet.Cells[i, 2] = linkTextModel;

                                            //if not have engine
                                            if (linkTextEngine.Contains(">"))
                                            {
                                                //write engine
                                                MySheet.Cells[i, 4] = linkTextEngine;

                                                //check Element exists
                                                CheckElementExist(i, MySheet);
                                                
                                                i++;
                                            }
                                            else
                                            {
                                                //get range year
                                                List<string> AllRangeYearLinks = GetArrayLinks();

                                                #region Range Year
                                                foreach (string linkTextRangeYear in AllRangeYearLinks)
                                                {
                                                    if (String.IsNullOrEmpty(linkTextRangeYear) || String.IsNullOrWhiteSpace(linkTextRangeYear))
                                                    {
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        string foldernameRangeYear = Path.Combine(foldernameEngine, linkTextRangeYear);
                                                        Directory.CreateDirectory(foldernameRangeYear);
                                                        //click engine
                                                        IWebElement elementRangeYear = PropertiesCollection.driver.FindElement(By.LinkText(linkTextRangeYear));

                                                        elementRangeYear.Click();

                                                        //write engine
                                                        MySheet.Cells[i, 3] = linkTextEngine;
                                                        //write range year
                                                        MySheet.Cells[i, 4] = linkTextRangeYear;

                                                        //check Element exists
                                                        CheckElementExist(i, MySheet);
                                                        
                                                        i++;
                                                        PropertiesCollection.driver.Navigate().Back();
                                                    }

                                                }
                                                #endregion
                                            }
                                            PropertiesCollection.driver.Navigate().Back();
                                        }

                                    }
                                    #endregion
                                    PropertiesCollection.driver.Navigate().Back();
                                }
                            }
                            PropertiesCollection.driver.Navigate().Back();
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            MyBook.SaveAs(@"d:output.xlsx");
            MyBook.Close();
            MyApp.Quit();
        }

        public static void GetDataForMake()
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
            MySheet.Cells[1, 1] = "MAKE";
            MySheet.Cells[1, 2] = "MODEL";
            MySheet.Cells[1, 3] = "ENGINE";
            MySheet.Cells[1, 4] = "RANGE YEAR";
            MySheet.Cells[1, 5] = "KEY PROGRAMMER";
            MySheet.Cells[1, 6] = "Tittle";
            MySheet.Cells[1, 7] = "IMG ";
            MySheet.Cells[1, 8] = "IMG 2";

            int i = 2;
            //click make
            IWebElement elementMake = PropertiesCollection.driver.FindElement(By.LinkText(make));
            elementMake.Click();

            //get models
            List<string> AllModelLinks = GetArrayLinks();

            foreach (string linkTextModel in AllModelLinks)
            {
                if (String.IsNullOrEmpty(linkTextModel) || String.IsNullOrWhiteSpace(linkTextModel))
                {
                    continue;
                }
                else
                {
                    //click model
                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.LinkText(linkTextModel));
                    elementModel.Click();

                    //get engines
                    List<string> AllEnginesLinks = GetArrayLinks();
                    #region Engine or Not
                    foreach (string linkTextEngine in AllEnginesLinks)
                    {
                        if (String.IsNullOrEmpty(linkTextEngine) || String.IsNullOrWhiteSpace(linkTextEngine))
                        {
                            continue;
                        }
                        else
                        {
                            //click engine
                            IWebElement elementEngine = PropertiesCollection.driver.FindElement(By.LinkText(linkTextEngine));
                            elementEngine.Click();

                            //if not have engine
                            if (linkTextEngine.Contains(">"))
                            {
                                //write make
                                MySheet.Cells[i, 1] = make;
                                //write model
                                MySheet.Cells[i, 2] = linkTextModel;
                                //write engine
                                MySheet.Cells[i, 4] = linkTextEngine;

                                //check Element exists
                                CheckElementExist(i, MySheet);

                                i++;
                            }
                            else
                            {
                                //get range year
                                List<string> AllRangeYearLinks = GetArrayLinks();

                                #region Range Year
                                foreach (string linkTextRangeYear in AllRangeYearLinks)
                                {
                                    if (String.IsNullOrEmpty(linkTextRangeYear) || String.IsNullOrWhiteSpace(linkTextRangeYear))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        //click engine
                                        IWebElement elementRangeYear = PropertiesCollection.driver.FindElement(By.LinkText(linkTextRangeYear));

                                        elementRangeYear.Click();

                                        //write make
                                        MySheet.Cells[i, 1] = make;
                                        //write model
                                        MySheet.Cells[i, 2] = linkTextModel;
                                        //write engine
                                        MySheet.Cells[i, 3] = linkTextEngine;
                                        //write range year
                                        MySheet.Cells[i, 4] = linkTextRangeYear;

                                        //check Element exists
                                        CheckElementExist(i, MySheet);
                                        i++;
                                        PropertiesCollection.driver.Navigate().Back();
                                    }

                                }
                                #endregion
                            }
                            PropertiesCollection.driver.Navigate().Back();
                        }

                    }
                    #endregion
                    PropertiesCollection.driver.Navigate().Back();
                }
            }
            string path = Environment.CurrentDirectory;
            string foldername = Path.Combine(path, "output.xlsx");
            //Directory.CreateDirectory(foldername);

            MyBook.SaveAs(foldername);
            MyBook.Close();
            MyApp.Quit();
            //PropertiesCollection.driver.Navigate().Back();
        }


        public static void SaveImage(string foldername, IWebElement elementimgSource, string linkTextMake, string linkTextModel, string linkTextEngine, string linkTextRangeYear, int countImg)
        {

            string nameofimg = linkTextMake + "_" + linkTextModel + "_" + linkTextEngine + "_" + linkTextRangeYear + "_" + countImg.ToString() + ".jpg";
            nameofimg = nameofimg.Replace("/", "-").Replace(">", "-");


            string filename = Path.Combine(foldername, nameofimg);
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(elementimgSource.GetAttribute("src"), filename);
            }
        }
        public static IWebElement waitForPageUntilElementIsVisible(By locator, int maxSeconds)
        {
            return new WebDriverWait(PropertiesCollection.driver, TimeSpan.FromSeconds(maxSeconds)).Until(ExpectedConditions.ElementExists((locator)));
        }
        public static void CheckElementExist(int i, Excel.Worksheet MySheet)
        {
            //check key program exists
            //bool isKeyDisplayed = PropertiesCollection.driver.FindElement(By.ClassName("WordSection1")).Displayed;
            By byKey = By.ClassName("WordSection1");
            var elementKeyProgram = PropertiesCollection.driver.FindElements(byKey).Count >= 1 ? PropertiesCollection.driver.FindElement(byKey) : null;
            if (elementKeyProgram != null)
            {
                //IWebElement elementkeyProgram = PropertiesCollection.driver.FindElement(By.ClassName("WordSection1"));

                //write key
                MySheet.Cells[i, 5] = elementKeyProgram.Text;
            }

            int colImg = 7;
            //check img exists
            //bool isImgDisplayed = PropertiesCollection.driver.FindElement(By.ClassName("locatorimage")).Displayed;
            By byImg = By.ClassName("locatorimage");
            var elementImgSource = PropertiesCollection.driver.FindElements(byImg).Count >= 1 ? PropertiesCollection.driver.FindElement(byImg) : null;
            if (elementImgSource != null)
            {

                IList<IWebElement> listImg = PropertiesCollection.driver.FindElements(By.ClassName("locatorimage"));
                foreach (IWebElement elementimgSource in listImg)
                {
                    MySheet.Cells[i, colImg] = "\n" + elementimgSource.GetAttribute("src");
                    colImg++;
                }

            }

            //write tittle
            By byTittle = By.XPath("/html/body/div[1]/div[2]/div[1]/p");
            var elementTittle = PropertiesCollection.driver.FindElements(byTittle).Count >= 1 ? PropertiesCollection.driver.FindElement(byTittle) : null;
            //bool isTittleisplayed = PropertiesCollection.driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div[1]/p")).Displayed;
            if (elementTittle != null)
            {
                //IWebElement elementtittle = PropertiesCollection.driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div[1]/p"));
                MySheet.Cells[i, 6] = elementTittle.Text;
            }
        }
        public static List<string> GetArrayLinks()
        {
            List<string> matchingLinks = new List<string>();

            ReadOnlyCollection<IWebElement> linksAllMake = PropertiesCollection.driver.FindElements(By.TagName("a"));

            foreach (IWebElement link in linksAllMake)
            {
                string text = link.Text;
                matchingLinks.Add(text);
            }

            return matchingLinks;
        }
        public static List<string> GetArrayLinks2()
        {
            List<string> matchingLinks = new List<string>();

            ReadOnlyCollection<IWebElement> linksAllMake = PropertiesCollection.driver.FindElements(By.Id("flsMakeSelect"));

            foreach (IWebElement link in linksAllMake)
            {
                string text = link.Text;
                matchingLinks.Add(text);
            }

            return matchingLinks;
           
        }
    }
}




