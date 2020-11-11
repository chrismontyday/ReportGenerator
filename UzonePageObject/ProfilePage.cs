using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using USFS.Library.TestAutomation;
using Serilog;
using System.Threading;
using USFS.Library.TestAutomation.Util;

namespace UzonePageObject
{
    public class ProfilePage : BasePage
    {
        public UtilityHelper util = new UtilityHelper();

        public void GetTeamMemberPhoto(TeamMember tm)
        {
            By search = By.XPath("//input[@id='txtSearchTerm']");
            By profileLink = By.XPath("//a[contains(text(),'" + tm.Name.Trim() + "')]");
            By image = By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/sk-profilepicture[1]");
            By profilePicture = By.XPath("//body/img[1]");
            By noProfile = By.XPath("//td[contains(text(),'Please enter a search term, or select all users fr')]");
            
            try
            {
                Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/user/index.html#/all");
                BrowserUtils.WaitForDisplayed(search, 30);
                Driver.FindElement(search).Click();
                Driver.FindElement(search).SendKeys(tm.Name);
                Driver.FindElement(search).SendKeys(Keys.Enter);

                //If Team Member's profile can not be found via search this changes their name to null so they can be removed from the list. 
                if (BrowserUtils.WaitForDisplayed(profileLink, 30) || !BrowserUtils.WaitForDisplayed(noProfile, 30))
                {
                    Log.Information("Team member profile pic found - " + tm.Name.ToString());
                    Driver.FindElement(profileLink).Click();
                    BrowserUtils.WaitForDisplayed(image, 15);
                    tm.PhotoPath = Driver.FindElement(image).GetAttribute("img");
                    Driver.Navigate().GoToUrl(@"https://uzone.unitedshore.com/user" + tm.PhotoPath);
                    IWebElement photo = Driver.FindElement(profilePicture);
                    tm.PhotoFilePath = util.TakeScreenshotOfElement(Driver, photo, tm.Name, tm.Location, tm.Id, true);
                    Log.Information("Screenshot of profile pic element taken successfully - \nProfile Pic Location: " + photo);
                }
                else
                {
                    Log.Information("Team member profile pic NOT found: " + tm.Name + " is set to null.");
                    tm.Name = null;
                }

            }
            catch (Exception e)
            {
                Log.Error(e, "Automation in GetTeamMemberPhoto() has failed");
                throw new Exception("GetTeamMemberPhoto Failed", e);
            }
        }
    }
}
