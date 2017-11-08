using GetDataOnWeb_v01;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KeyProgrammerProcedure_v01
{
    public partial class GetDataOnWebUI : Form
    {
        public GetDataOnWebUI()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        //Navigate button
        private void button1_Click(object sender, EventArgs e)
        {
            //build Key Programmer Procedure
            if (textBox1.Text.Contains("wikirke"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL("http://www.wikirke.co.uk/manufacturers.php?pt=1");

                //add make to combobox
                List<string> AllMakeLinks = KeyProgrammerProcedure.GetArrayLinks();
                foreach (string linkTextMake in AllMakeLinks)
                {
                    if (String.IsNullOrEmpty(linkTextMake) || String.IsNullOrWhiteSpace(linkTextMake))
                    {
                        continue;
                    }
                    else
                    {
                        cmbMakes.Items.Add(linkTextMake);
                    }
                }
            }

            //build Key Programmer Procedure
            if (textBox1.Text.Contains("owner"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL("https://owner.ford.com/tools/account/maintenance/maintenance-schedule.html#/details");

                //add make to combobox
                List<string> AllMakeLinks = KeyProgrammerProcedure.GetArrayLinks2();

                foreach (string linkTextMake in AllMakeLinks)
                {
                    if (String.IsNullOrEmpty(linkTextMake) || String.IsNullOrWhiteSpace(linkTextMake))
                    {
                        continue;
                    }
                    else
                    {
                        cmbMakes.Items.Add(linkTextMake);
                    }
                }
            }

            //build Schedule Maintenance
            if (textBox1.Text.Contains("edmunds"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL("http://edmunds.mashery.com/io-docs");

                ScheduleMaintence.GetAllYearMake("sv68rfxgdc7qxvc9payea3fq");

                for (int i = 0; i < ScheduleMaintence.arrayAllYearMake.Count(); i++)
                {
                    cmbMakes.Items.Add(ScheduleMaintence.arrayAllYearMake[i].Make);

                }
            }

            //build Battery Finder
            if (textBox1.Text.Contains("autobatteries"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL("http://www.autobatteries.com/en-us/car-battery-finder");

                IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("product_finder_year"));
                var options = elementYears.AsDropDown().Options;
                foreach (IWebElement linkTextYear in options)
                {
                    if (String.IsNullOrEmpty(linkTextYear.Text) || String.IsNullOrWhiteSpace(linkTextYear.Text))
                    {
                        continue;
                    }
                    else
                    {
                        cmbYears.Items.Add(linkTextYear.Text);
                    }
                }
            }

            //build Autocodes
            if (textBox1.Text.Contains("autocodes"))
            {
                //add list Make to combobox Make
                foreach (var item in Autocodes.listMake)
                {
                    cmbMakes.Items.Add(item);
                }
            }

            //build Source BCI
            if (textBox1.Text.Contains("sourcebci"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL("http://sourcebci.com/Account/Login.aspx?error=The%20user%20session%20has%20timed%20out.%20%20Please%20log%20in%20again.");
                System.Threading.Thread.Sleep(5000);

                //Input username
                IWebElement InputUsername = PropertiesCollection.driver.FindElement(By.Id("MainContent_LoginUser_UserName"));
                InputUsername.SendKeys("Sonle");
                System.Threading.Thread.Sleep(1000);

                //Input Password
                IWebElement InputPassword = PropertiesCollection.driver.FindElement(By.Id("MainContent_LoginUser_Password"));
                InputPassword.SendKeys("Innova17352");
                System.Threading.Thread.Sleep(1000);

                //Click Login
                IWebElement findBtnLogin = PropertiesCollection.driver.FindElement(By.Id("MainContent_LoginUser_LoginButton"));
                findBtnLogin.Click();
                System.Threading.Thread.Sleep(5000);

                //Load page Passenger Car and light trucks
                PropertiesCollection.NavigatetoURL("http://sourcebci.com/app_passcartruck.aspx");
                System.Threading.Thread.Sleep(4000);

                //Sourcebci.GetDataAll();
                GetDataSourcebci.WriteDataExcel();
                //MessageBox.Show("DONE");
            }

            //build Possible Causes via wiki-cross
            if (textBox1.Text.Contains("wiki.ross"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL(textBox1.Text);

            }

            //build collect data of DTC code
            if (textBox1.Text.Contains("repair.alldata"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL("https://repair.alldata.com/alldata/secure/login.action");
            }
        }

        //get Img for Make/Year
        private void btnBuildData_Click(object sender, EventArgs e)
        {
            KeyProgrammerProcedure.make = cmbMakes.SelectedItem.ToString();
            KeyProgrammerProcedure.GetImgforMake();
            MessageBox.Show("DONE");
        }

        //get Data for Make/Year
        private void btnGetData_Click(object sender, EventArgs e)
        {
            //get data of Key Programmer Procedure
            if (textBox1.Text.Contains("wikirke"))
            {
                KeyProgrammerProcedure.make = cmbMakes.SelectedItem.ToString();
                KeyProgrammerProcedure.GetDataForMake();
                MessageBox.Show("DONE");
            }

            //get data of Schedule_Maintenance_Ford
            if (textBox1.Text.Contains("owner"))
            {
                SM_Ford.WriteDataExcel();
                MessageBox.Show("DONE");
            }

            //get data of Battery Finder
            if (textBox1.Text.Contains("autobatteries"))
            {
                BatteryFinder.year = cmbYears.SelectedItem.ToString();
                BatteryFinder.GetDataForYear();
                MessageBox.Show("DONE");
            }

            //get data of Schedule Maintenance
            if (textBox1.Text.Contains("edmunds"))
            {

                ScheduleMaintence.HandleonWebforSelectedMakeYear(cmbMakes.Text, cmbYears.Text);

                MessageBox.Show("DONE");
            }

            //get data of Autocodes by Make
            if (textBox1.Text.ToLower().Contains("autocodes"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL(textBox1.Text);
                Autocodes.GetDatabyMake(textBox1.Text);
                MessageBox.Show("DONE");
            }

            //get Possible Causes via wiki-cross for a page
            if (txtSubLink.Text.ToLower().Contains("wiki.ross"))
            {
                //navigate to URL
                PropertiesCollection.NavigatetoURL(txtSubLink.Text);
                PossibleCauseWikiRoss.GetPosCauseForPage();
                MessageBox.Show("DONE");
            }

            //Get database CCA
            if (textBox1.Text.Contains("sourcebci"))
            {
                if (String.IsNullOrEmpty(txtUsername.Text) || String.IsNullOrEmpty(txtPassword.Text))
                {
                    MessageBox.Show("Please enter username and password");
                }
                else
                {
                    CommonMethods.LogIntoWebsite(txtUsername.Text, txtPassword.Text, By.ClassName("textEntry"), By.ClassName("passwordEntry"), By.Name("ctl00$MainContent$LoginUser$LoginButton"));
                    System.Threading.Thread.Sleep(3000);

                    //Passenger Cars and Light Trucks
                    //IWebElement elementPassenger = PropertiesCollection.driver.FindElement(By.LinkText("Passenger Cars and Light Trucks"));//LinkText("Passenger Cars and Light Trucks"));
                    //elementPassenger.Click();

                    PropertiesCollection.NavigatetoURL("http://sourcebci.com/app_passcartruck.aspx");
                    System.Threading.Thread.Sleep(1000);

                    #region Add YMME To Combobox
                    //add make to combobox
                    //IWebElement elementMake = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddMake1"));
                    //var options_makes = elementMake.AsDropDown().Options;
                    //foreach (IWebElement linkTextMake in options_makes)
                    //{
                    //    if (String.IsNullOrEmpty(linkTextMake.Text) || String.IsNullOrWhiteSpace(linkTextMake.Text))
                    //    {
                    //        continue;
                    //    }
                    //    else
                    //    {
                    //        cmbMakes.Items.Add(linkTextMake.Text);

                    //        if (cmbMakes != null)
                    //        {
                    //            //add years to combobox
                    //            IWebElement elementYears = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddYear1"));
                    //            var options = elementYears.AsDropDown().Options;
                    //            foreach (IWebElement linkTextYear in options)
                    //            {
                    //                if (String.IsNullOrEmpty(linkTextYear.Text) || String.IsNullOrWhiteSpace(linkTextYear.Text))
                    //                {
                    //                    continue;
                    //                }
                    //                else
                    //                {
                    //                    cmbYears.Items.Add(linkTextYear.Text);
                    //                }
                    //            }
                    //        }
                    //    }
                    //}

                    //add model to combobox
                    IWebElement elementModel = PropertiesCollection.driver.FindElement(By.Id("MainContent_ddModel1"));
                    var options_model = elementModel.AsDropDown().Options;
                    foreach (IWebElement linkTextModel in options_model)
                    {
                        if (String.IsNullOrEmpty(linkTextModel.Text) || String.IsNullOrWhiteSpace(linkTextModel.Text))
                        {
                            continue;
                        }
                        else
                        {
                            cmbModel.Items.Add(linkTextModel.Text);
                        }
                    }
                    #endregion

                    //Sourcebci.GetDataAll();
                    GetDataSourcebci.WriteDataExcel();
                    MessageBox.Show("DONE");
                }
            }
            if (textBox1.Text.Contains("repair.alldata"))
            {
                if (String.IsNullOrEmpty(txtUsername.Text) || String.IsNullOrEmpty(txtPassword.Text))
                {
                    MessageBox.Show("Please enter username and password");
                }
                else
                {
                    CommonMethods.LogIntoWebsite(txtUsername.Text, txtPassword.Text, By.Id("login"), By.Id("password"), By.ClassName("login_shrt_btn_rdy"));
                    System.Threading.Thread.Sleep(3000);
                   
                    //username: vninnova -- --- password: conghoa364
                    PropertiesCollection.NavigatetoURL("https://repair.alldata.com/alldata/secure/login.action");
                    System.Threading.Thread.Sleep(2000);
                   
                    GetdataDTC.GetAllDataDTC();
                    MessageBox.Show("DONE");
                }               
            }
            Console.Clear();
        }

        //get Data for all Makes/Years
        private void btnGetAll_Click(object sender, EventArgs e)
        {
            //get data of Battery Finder
            if (textBox1.Text.Contains("autobatteries"))
            {

                BatteryFinder.GetDataAll();
                MessageBox.Show("DONE");
            }

            //get data of Source BCI
            if (textBox1.Text.Contains("sourcebci"))
            {
                if (String.IsNullOrEmpty(txtUsername.Text) || String.IsNullOrEmpty(txtPassword.Text))
                {
                    MessageBox.Show("Please enter username and password");
                }
                else
                {
                    CommonMethods.LogIntoWebsite(txtUsername.Text, txtPassword.Text, By.ClassName("textEntry"), By.ClassName("passwordEntry"), By.Name("ctl00$MainContent$LoginUser$LoginButton"));
                    System.Threading.Thread.Sleep(1000);

                    //Passenger Cars and Light Trucks
                    //IWebElement elementPassenger = PropertiesCollection.driver.FindElement(By.LinkText("Passenger Cars and Light Trucks"));
                    //elementPassenger.Click();
                    PropertiesCollection.NavigatetoURL("http://sourcebci.com/app_passcartruck.aspx");
                    System.Threading.Thread.Sleep(1000);

                    Sourcebci.GetDataAll();

                    MessageBox.Show("DONE");
                }

            }

            //get data of Schedule Maintenance
            if (textBox1.Text.Contains("edmunds"))
            {

                ScheduleMaintence.HandleonWeb();

                MessageBox.Show("DONE");
            }
        }

        private void cmbMakes_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbMakes_SelectedValueChanged(object sender, EventArgs e)
        {
            //Schedule Maintenance task
            if (textBox1.Text.Contains("edmunds"))
            {
                cmbYears.Items.Clear();
                for (int i = 0; i < ScheduleMaintence.arrayAllYearMake.Count(); i++)
                {
                    for (int j = 0; j < ScheduleMaintence.arrayAllYearMake[i].Year.Count(); j++)
                    {
                        if (cmbMakes.SelectedItem.ToString().Contains(ScheduleMaintence.arrayAllYearMake[i].Make))
                        {
                            cmbYears.Items.Add(ScheduleMaintence.arrayAllYearMake[i].Year[j].Year);
                        }
                    }
                }
            }

            //AutoCodes task
            if (textBox1.Text.ToLower().Contains("autocodes"))
            {
                Autocodes.make = cmbMakes.SelectedItem.ToString();
            }
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void cmbYears_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
