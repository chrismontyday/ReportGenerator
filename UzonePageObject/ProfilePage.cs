using OpenQA.Selenium;
using System;
using USFS.Library.TestAutomation;
using Serilog;
using USFS.Library.TestAutomation.Util;
using System.Threading;

namespace UzonePageObject
{
    public class ProfilePage : BasePage
    {
        public UtilityHelper util = new UtilityHelper();

        public void GetTeamMemberPhoto(TeamMember tm)
        {
            By search = By.XPath("//input[@id='txtSearchTerm']");
            By noProfile = By.XPath("//td[contains(text(),'Please enter a search term, or select all users from the drop down.')]");
            By image = By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/sk-profilepicture[1]");
            By profilePicture = By.XPath("//body/img[1]");
            By individualLink = By.XPath(@"//a[contains(text(),'" + tm.Name + "')]");
            By profileLink = By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/table[1]/tbody[1]/tr[1]/td[2]/strong[1]/a[1]");

            Log.Information(tm.Name + " GetTeamMemberPhoto() has started.");

            try
            {
                int threeTries = 3;
                do
                {
                    Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/user/index.html#/all");
                    BrowserUtils.WaitForDisplayed(search, 90);
                    Driver.FindElement(search).Click();
                    Driver.FindElement(search).Clear();
                    Driver.FindElement(search).SendKeys(tm.Name);
                    Driver.FindElement(search).SendKeys(Keys.Enter);

                    if (BrowserUtils.WaitForDisplayed(individualLink, 90) || !BrowserUtils.WaitForDisplayed(noProfile, 90))
                    {
                        Driver.FindElement(profileLink).Click();
                        BrowserUtils.WaitForDisplayed(image, 90);
                        tm.PhotoPath = Driver.FindElement(image).GetAttribute("img");
                        Driver.Navigate().GoToUrl(@"https://uzone.unitedshore.com/user" + tm.PhotoPath);
                        IWebElement photo = Driver.FindElement(profilePicture);
                        tm.PhotoFilePath = util.TakeScreenshotOfElement(Driver, photo, tm.Name, tm.Id, true);
                        Log.Information(tm.Name + " Screenshot of profile pic element taken successfully");
                        break;
                    }

                } while (threeTries-- != 0 && tm.PhotoFilePath == null);

                //if the leader's profile pic could not be found this is where it notes that and skips them. 
                if (tm.PhotoFilePath == null)
                {                    
                    tm.Skipped = true;
                    Log.Information(tm.Name + " was skipped.");
                    tm.SkippedNote = tm.Name = " was skipped because bot could not find profile.";
                }
            }
            catch (Exception e)
            {                
                throw new Exception("GetTeamMemberPhoto() Method Failed", e);
            }
        }
    }
}
