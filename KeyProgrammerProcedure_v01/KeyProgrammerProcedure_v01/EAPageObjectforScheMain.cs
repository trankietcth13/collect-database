using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetDataOnWeb_v01
{
    class EAPageObject
    {
        public EAPageObject()
        {
            PageFactory.InitElements(PropertiesCollection.driver, this);
        }

        [FindsBy(How = How.XPath, Using = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/table/tbody/tr[1]/td[2]/select")]
        public IWebElement ddStateforYearMakeId { get; set; }

        [FindsBy(How = How.Id, Using = "apiKey")]
        public IWebElement txtAPIKey { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/div/span[1]")]
        public IWebElement btnGetAllYearsMakes { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/input[5]")]
        public IWebElement btnTryitAllYearsMakes { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/table/tbody/tr[2]/td[2]/input")]
        public IWebElement txtYear { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/div[1]/div[4]/div[1]/ul[5]/li[1]/ul/li[1]/form/div[1]/pre[1]")]
        public IWebElement RequestURL { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/pre")]
        public IWebElement JsonAllText { get; set; }
    }
}
