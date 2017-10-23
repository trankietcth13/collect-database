using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Table;
using System.Linq.Expressions;
using OpenQA.Selenium.Chrome;

namespace GetDataOnWeb_v01
{
    public static class CommonMethods
    {
        public static void SwitchToTab(string pageId)
        {
            
            PropertiesCollection.driver.SwitchTo().Window(pageId);
        }

        public static void DoWait(int milliseconds)
        {
            var wait = new WebDriverWait(PropertiesCollection.driver, TimeSpan.FromMilliseconds(milliseconds));
            var waitComplete = wait.Until<bool>(
                arg =>
                {
                    System.Threading.Thread.Sleep(milliseconds);                    
                    return true;
                });
        }
      

        //list string from Ilist
        public static List<string> ListStringfromIList(IList<IWebElement> listElements)
        {
            List<string> matchingLinks = new List<string>(listElements.Count);
           // System.Threading.Thread.Sleep(5000);
            foreach (IWebElement linkTextElement in listElements)
            {                
               matchingLinks.Add(linkTextElement.Text);                                   
            }
           // System.Threading.Thread.Sleep(5000);
            return matchingLinks;
        }


        //list string from Elements have the same name value
        public static List<string> GetListfromtheSameElements( By by)
        {
            List<string> matchingLinks = new List<string>();

            ReadOnlyCollection<IWebElement> linksAllElements = PropertiesCollection.driver.FindElements(by);
            System.Threading.Thread.Sleep(5000);
            foreach (IWebElement element in linksAllElements)
            {
                string text = element.Text;
                matchingLinks.Add(text);

            }
            System.Threading.Thread.Sleep(5000);
            return matchingLinks;
        }

        //array string from Elements have the same name value
        public static string[] GetArrayfromtheSameElements(By by)
        {
            ReadOnlyCollection<IWebElement> linksAllElements = PropertiesCollection.driver.FindElements(by);

            string[] matchingLinks = new string[linksAllElements.Count];
            int i = 0;
            foreach (IWebElement element in linksAllElements)
            {
                string text = element.Text;
                matchingLinks[i] = text;
                i++;
            }
            System.Threading.Thread.Sleep(5000);
            return matchingLinks;
        }

        //check element exist
        public static bool IsElementPresent(By by)
        {
            try
            {
                PropertiesCollection.driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        //get al option of drop down list
        public static SelectElement AsDropDown(this IWebElement webElement)
        {
            return new SelectElement(webElement);
        }

        //private Stream CreateExcelFile(Stream stream = null, List<string> nameofallSheet )
        //{
        //    var list = CreateTestItems();
        //    using (var excelPackage = new ExcelPackage(stream ?? new MemoryStream()))
        //    {
        //        // Tạo author cho file Excel
        //        //excelPackage.Workbook.Properties.Author = "Hanker";
        //        // Tạo title cho file Excel
        //        //excelPackage.Workbook.Properties.Title = "EPP test background";
        //        // thêm tí comments vào làm màu 
        //        //excelPackage.Workbook.Properties.Comments = "This is my fucking generated Comments";
        //        // Add Sheet vào file Excel
        //        excelPackage.Workbook.Worksheets.Add("First Sheet");
        //        // Lấy Sheet bạn vừa mới tạo ra để thao tác 
        //        var workSheet = excelPackage.Workbook.Worksheets[1];
        //        // Đổ data vào Excel file
        //        workSheet.Cells[1, 1].LoadFromCollection(list, true, TableStyles.Dark9);
        //        // BindingFormatForExcel(workSheet, list);
        //        excelPackage.Save();
        //        return excelPackage.Stream;
        //    }
        //}
        //log-in
        public static void LogIntoWebsite(string username, string password, By byUsername, By byPassword, By byLoginButton)
        {
            //enter username from txtUsername
            IWebElement FindUsername = PropertiesCollection.driver.FindElement(byUsername);
            FindUsername.SendKeys(username);

            System.Threading.Thread.Sleep(1000);

            //enter password from txtPassword
            IWebElement FindPassword = PropertiesCollection.driver.FindElement(byPassword);
            FindPassword.SendKeys(password);

            System.Threading.Thread.Sleep(1000);

            //click log-in button
            IWebElement FindLoginButton = PropertiesCollection.driver.FindElement(byLoginButton);
            FindLoginButton.Click();

            System.Threading.Thread.Sleep(1000);
        }

        public static void SwitchToWindow(Expression<Func<IWebDriver, bool>> predicateExp)
        {
            var predicate = predicateExp.Compile();
            foreach (var handle in PropertiesCollection.driver.WindowHandles)
            {
                PropertiesCollection.driver.SwitchTo().Window(handle);
                if (predicate(PropertiesCollection.driver))
                {
                    return;
                }
            }

            throw new ArgumentException(string.Format("Unable to find window with condition: '{0}'", predicateExp.Body));
        }

    }
}
