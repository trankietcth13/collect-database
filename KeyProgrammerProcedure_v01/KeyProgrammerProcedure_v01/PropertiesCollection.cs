using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataOnWeb_v01
{
    enum PropertyType
    {
        Id,
        Name,
        LinkText,
        CssName,
        ClassName,
        XPath
    }
    enum MakeForKeyProgrammer
    {
        ACURA,
        AUDI,
        BMW,
        BUICK,
        CADILLAC,
        CHEVROLET,
        FORD,
        GEO,
        GMC,
        HONDA,
        HUMMER,
        HYUNDAI,
        INFINITY,
        KIA,
        LEXUS,
        LINCOLN,
        MAZDA,
        MERCEDES,
        MERCURY,
        MITSUBISHI,
        NISSAN,
        OLDSMOBILE,
        PONTIAC,
        SUZUKI,
        TOYOTA,
        VW
    }
    class PropertiesCollection
    {

        //Auto-implemented Property
        public static IWebDriver driver { get; set; }

        //driver for switch tab
        //public static IWebDriver driverSTab = new ChromeDriver();
        //open website
        public static void NavigatetoURL(string url)
        {
            //Navigate to Web
            PropertiesCollection.driver.Navigate().GoToUrl(url);
            Console.WriteLine("Opened URL");

        }

       
    }
}
