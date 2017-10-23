using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataOnWeb_v01
{
    class GetMethods
    {

        public static string GetText(IWebElement element)
        {
            return element.GetAttribute("value");
        }
        public static string GetTextFromDropDown(IWebElement element)
        {
            return new SelectElement(element).AllSelectedOptions.SingleOrDefault().Text;
        }

    }
}
