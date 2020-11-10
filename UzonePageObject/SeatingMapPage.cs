using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using USFS.Library.TestAutomation;
using Serilog;

namespace UzonePageObject
{
    public class SeatingMapPage : BasePage
    {
        //public IWebDriver Driver;
        public WebDriverWait Wait { get { return new WebDriverWait(Driver, TimeSpan.FromSeconds(60)); } }
        public UtilityHelper util = new UtilityHelper();

        //Takes screenshot of Team Member's seating location on seating map. 
        public void GetTeamMemberSeatingMap(TeamMember tm)
        {
            By searchBox = By.XPath("//span[@id='select2-chosen-1']");
            By searchSeatMap = By.XPath("//input[@id='s2id_autogen1_search']");
            By fullScreen = By.XPath("//body/sk-modalpresenter[1]/sk-modalpresentationlayer[1]");
            By fullScreenTwo = By.XPath("//div[@id='select2-drop-mask']");
            By noMatch = By.XPath("//li[contains(text(),'No matches found')]");
            By seatingMap = By.XPath("//div[@id='mapContainer']//div[@id='draggable']");
            By mapSelect = By.XPath("//select[@id='mapSelect']");

            try
            {
                Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/map/#/map/51");
                Wait.Until(d => d.FindElement(searchBox).Displayed);
                Driver.FindElement(searchBox).Click();
                Driver.FindElement(searchSeatMap).SendKeys(tm.Name);
                Driver.FindElement(searchSeatMap).SendKeys(Keys.Enter);

                if (WaitForVisible(noMatch) || !WaitForVisible(fullScreen))
                {
                    //If Team Member does not have personal seating map. This re-routes to generic map. 
                    Driver.Navigate().Refresh();
                    Wait.Until(d => d.FindElement(mapSelect));
                    Driver.FindElement(mapSelect).Click();
                    Driver.FindElement(mapSelect).SendKeys(tm.Location);
                    Driver.FindElement(mapSelect).SendKeys(Keys.Enter);
                }
                else
                {
                    Driver.FindElement(fullScreen).Click();
                }

                Wait.Until(d => d.FindElement(seatingMap).Displayed);
                IWebElement map = Driver.FindElement(seatingMap);
                tm.MapFilePath = util.TakeScreenshotOfElement(Driver, map, tm.Name, tm.Location, tm.Id, false);

            }
            catch (Exception)
            {
                Log.Information("Automation in GetTeamMemberSeatingMap() has failed");

            }
        }

        //Returns bool if web element does not become visible. 
        private bool WaitForVisible(By path)
        {
            bool isVisible = false;

            try
            {
                Driver.FindElement(path);
                isVisible = true;
            }
            catch (Exception)
            {
                isVisible = false;
            }

            return isVisible;
        }
    }
}
