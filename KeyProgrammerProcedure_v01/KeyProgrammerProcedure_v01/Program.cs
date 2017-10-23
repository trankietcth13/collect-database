
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetDataOnWeb_v01
{
    class Program
    {
       
        static void Main(string[] args)
        {

            PropertiesCollection.driver = new ChromeDriver();

            //Create the reference for browser
            //IWebDriver driver = new ChromeDriver();
            
            Application.Run(new KeyProgrammerProcedure_v01.GetDataOnWebUI());
            PropertiesCollection.driver.Close();
            Console.WriteLine("Close");
            
        }

       
        //public void Initialize()
        //{
        //    //Navigate to Web
        //    driver.Navigate().GoToUrl("http://edmunds.mashery.com/io-docs");
        //    Console.WriteLine("Opened URL");
        //}

        
        //public void ExecuteTest()
        //{
        //    //Find the Element
        //    IWebElement element = driver.FindElement(By.Name("key"));

        //    //Perform Ops
        //    element.SendKeys("sv68rfxgdc7qxvc9payea3fq");

        //    Console.WriteLine("Execute");
        //}

        
        //public void NextTest()
        //{
        //    Console.WriteLine("Next method");
        //}
        
        //public void CleanUp()
        //{
        //    driver.Close();
        //    Console.WriteLine("Close");
        //}
    }
}
