using OpenQA.Selenium;
using System;
using USFS.Library.TestAutomation;
using Serilog;
using USFS.Library.TestAutomation.Util;
using System.Threading;
using OpenQA.Selenium.Support.UI;

namespace UzonePageObject
{
    public class ProfilePage : BasePage
    {
       // public WebDriverWait Wait { get { return new WebDriverWait(Driver, TimeSpan.FromSeconds(90)); } }
        public UtilityHelper util = new UtilityHelper();

        public void GetTeamMemberPhoto(TeamMember tm)
        {
            By search = By.XPath("//input[@id='txtSearchTerm']");
            By individualLink = By.XPath(@"//a[contains(text(),'" + tm.Name + "')]");
            By noProfile = By.XPath("//td[contains(text(),'Please enter a search term, or select all users from the drop down.')]");
            By image = By.XPath("//body/div[@id='js-body-content']/div[@id='js-content-area-full']/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/sk-profilepicture[1]");
            By profilePicture = By.XPath("//body/img[1]");

            try
            {
                int threeTries = 3;
                do
                {
                    Driver.Navigate().GoToUrl("https://www.uzone.uwm.com/user/index.html#/all");
                    BrowserUtils.WaitForDisplayed(search, 10);
                    Driver.FindElement(search).Click();
                    Driver.FindElement(search).Clear();
                    Driver.FindElement(search).SendKeys(tm.Name);
                    Driver.FindElement(search).SendKeys(Keys.Enter);
                    BrowserUtils.WaitFor(1);

                    if (BrowserUtils.WaitForDisplayed(individualLink, 15) || !BrowserUtils.WaitForDisplayed(noProfile, 10))
                    {
                        Driver.FindElement(individualLink).Click();
                        BrowserUtils.WaitForDisplayed(image, 10);
                        tm.PhotoPath = Driver.FindElement(image).GetAttribute("img");
                        Driver.Navigate().GoToUrl(@"https://www.uzone.uwm.com/user" + tm.PhotoPath);
                        IWebElement photo = Driver.FindElement(profilePicture);
                        tm.PhotoFilePath = util.TakeScreenshotOfElement(Driver, photo, tm.Name, tm.Id, true);
                        Log.Information(tm.Name + " Screenshot of profile picture web element taken successfully");
                        break;
                    }

                } while (threeTries-- != 0);

                //if the leader's profile pic could not be found this is where it notes that and skips them. 
                if (tm.PhotoFilePath == null)
                {                    
                    tm.Skipped = true;
                    Log.Information(tm.Name + " was SKIPPED! ");                    
                    tm.SkippedNote = tm.Name + " - Skipped because bot could not find profile. ";
                }
            }
            catch (Exception e)
            {                
                throw new Exception("GetTeamMemberPhoto() Method Failed", e);
            }
        }
    }
}
