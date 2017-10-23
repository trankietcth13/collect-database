using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataOnWeb_v01
{
    class EAPageObjectforKeyProgram
    {
        public EAPageObjectforKeyProgram()
        {
            PageFactory.InitElements(PropertiesCollection.driver, this);
        }

        [FindsBy(How = How.TagName, Using = "a")]
        public IWebElement linkMake { get; set; }
    }
}
