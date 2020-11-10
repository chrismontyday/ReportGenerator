using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;

namespace UzonePageObject
{
    public class SeatingMapPage 
    {
        public IWebDriver WebDriver;
        public WebDriverWait Wait { get { return new WebDriverWait(WebDriver, TimeSpan.FromSeconds(60)); } }
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
                WebDriver.Navigate().GoToUrl("https://uzone.unitedshore.com/map/#/map/51");
                Wait.Until(d => d.FindElement(searchBox).Displayed);
                WebDriver.FindElement(searchBox).Click();
                WebDriver.FindElement(searchSeatMap).SendKeys(tm.Name);
                WebDriver.FindElement(searchSeatMap).SendKeys(Keys.Enter);

                if (WaitForVisible(noMatch) || !WaitForVisible(fullScreen))
                {
                    //If Team Member does not have personal seating map. This re-routes to generic map. 
                    WebDriver.Navigate().Refresh();
                    Wait.Until(d => d.FindElement(mapSelect));
                    WebDriver.FindElement(mapSelect).Click();
                    WebDriver.FindElement(mapSelect).SendKeys(tm.Location);
                    WebDriver.FindElement(mapSelect).SendKeys(Keys.Enter);
                }
                else
                {
                    WebDriver.FindElement(fullScreen).Click();
                }

                Wait.Until(d => d.FindElement(seatingMap).Displayed);
                IWebElement map = WebDriver.FindElement(seatingMap);
                tm.MapFilePath = util.TakeScreenshotOfElement(WebDriver, map, tm.Name, tm.Location, tm.Id, false);

            }
            catch (Exception)
            {
                Console.WriteLine("Failure occured in GetTeamMemberSeatingMap()"); ;
            }
        }

        //Returns bool if web element does not become visible. 
        private bool WaitForVisible(By path)
        {
            bool isVisible = false;

            try
            {
                WebDriver.FindElement(path);
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
