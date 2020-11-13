using OpenQA.Selenium;
using System;
using USFS.Library.TestAutomation;
using Serilog;
using USFS.Library.TestAutomation.Util;

namespace UzonePageObject
{
    public class SeatingMapPage : BasePage
    {
        public UtilityHelper util = new UtilityHelper();
        By searchBox = By.XPath("//span[@id='select2-chosen-1']");
        By searchSeatMap = By.XPath("//input[@id='s2id_autogen1_search']");
        By fullScreen = By.XPath("//body/sk-modalpresenter[1]/sk-modalpresentationlayer[1]");
        By seatingMap = By.XPath("//div[@id='mapContainer']//div[@id='draggable']");
        By mapSelect = By.XPath("//select[@id='mapSelect']");

        public void GetTeamMemberSeatingMap(TeamMember tm)
        {
            try
            {
                Log.Information("GetTeamMemberSeatingMap() has started");
                Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/map/#/map/51");
                BrowserUtils.WaitForDisplayed(searchBox, 30);
                Driver.FindElement(searchBox).Click();
                Driver.FindElement(searchSeatMap).SendKeys(tm.Name);
                Driver.FindElement(searchSeatMap).SendKeys(Keys.Enter);

                if (BrowserUtils.WaitForVisible(fullScreen, 25))
                {
                    Log.Information("Seating map for " + tm.Name.ToString() + " found!");
                    Driver.FindElement(fullScreen).Click();
                }
                else
                {
                    //If Team Member does not have personal seating map. This re-routes to generic map. 
                    //Log.Information("Seating Map does not exist for " + tm.Name.ToString() + ". Rerouting to generic map...");
                    //Driver.Navigate().Refresh();
                    //BrowserUtils.WaitForDisplayed(mapSelect, 30);
                    //Driver.FindElement(mapSelect).Click();
                    //Driver.FindElement(mapSelect).SendKeys(tm.Location);
                    //Driver.FindElement(mapSelect).SendKeys(Keys.Enter);
                }

                BrowserUtils.WaitForDisplayed(seatingMap, 30);
                IWebElement map = Driver.FindElement(seatingMap);
                tm.MapFilePath = util.TakeScreenshotOfElement(Driver, map, tm.Name, tm.Event, tm.Id, false);
                Log.Information("Screenshot of Seating Map for " + tm.Name.ToString() + " taken." +
                "\nSeating Map Location: " + tm.MapFilePath.ToString());
            }
            catch (Exception e)
            {
                Log.Error(e, "GetTeamMemberSeatingMap() has failed");
                throw new Exception("GetTeamMemberSeatingMap Failed", e);
            }
        }
    }
}
