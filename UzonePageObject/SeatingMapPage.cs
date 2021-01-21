using OpenQA.Selenium;
using System;
using USFS.Library.TestAutomation;
using Serilog;
using USFS.Library.TestAutomation.Util;
using OpenQA.Selenium.Interactions;

namespace UzonePageObject
{
    public class SeatingMapPage : BasePage
    {
        public UtilityHelper util = new UtilityHelper();
        By searchBox = By.XPath("//span[@id='select2-chosen-1']");
        By searchSeatMap = By.XPath("//input[@id='s2id_autogen1_search']");
        By fullScreen = By.XPath("//body/sk-modalpresenter[1]/sk-modalpresentationlayer[1]");
        By seatingMap = By.XPath("//div[@id='mapContainer']//div[@id='draggable']");
        By floorNumber = By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/h3[1]");
        By mapSelect = By.XPath("//select[@id='mapSelect']");

        public void GetTeamMemberSeatingMap(TeamMember tm)
        {
            try
            {
                int threeTries = 3;
                do
                {
                    Log.Information("Getting " + tm.Name + " Seating map");
                    Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/map/#/map/51");
                    BrowserUtils.WaitForDisplayed(searchBox, 90);
                    Driver.FindElement(searchBox).Click();
                    Driver.FindElement(searchSeatMap).SendKeys(tm.Name);
                    Driver.FindElement(searchSeatMap).SendKeys(Keys.Enter);

                    if (BrowserUtils.WaitForVisible(fullScreen, 90))
                    {
                        Actions action = new Actions(Driver);
                        action.MoveByOffset(0, 0).Click().Perform();
                    }

                    if (BrowserUtils.WaitForVisible(seatingMap, 90))
                    {
                        IWebElement map = Driver.FindElement(seatingMap);
                        tm.MapFilePath = util.TakeScreenshotOfElement(Driver, map, tm.Name, tm.Id, false);
                        tm.Floor = Driver.FindElement(floorNumber).Text;
                        break;
                    }

                } while (threeTries-- != 0 && tm.PhotoFilePath == null);

                if (tm.MapFilePath == null)
                {
                    Log.Information("Seating map failed for " + tm.Name);
                    tm.SkippedNote = tm.Name + " was skipped because bot could not find their seating map.";
                }
                else
                {
                    tm.Skipped = true;
                    Log.Information("Screenshot of Seating Map for " + tm.Name + " taken successfully.");
                }
            }
            catch (Exception e)
            {
                throw new Exception("GetTeamMemberSeatingMap Failed", e);
            }
        }
    }
}
