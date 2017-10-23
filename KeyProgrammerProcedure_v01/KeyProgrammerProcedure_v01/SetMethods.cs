using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Support.UI;

namespace GetDataOnWeb_v01
{
    class SetMethods
    {
        //Enter Text
        public static void EnterText( IWebElement element, string value)
        {
            element.SendKeys(value);
        }

        //Click button, checkbox, option etc
        public static void Click( IWebElement element)
        {
            element.Click();
        }

        //Selecting a drop down
        public static void SelectDropDown( IWebElement element, string value)
        {
            
           new SelectElement(element).SelectByText(value);

        }
    }
}
