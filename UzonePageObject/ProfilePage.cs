using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using USFS.Library.TestAutomation;
using Serilog;

namespace UzonePageObject
{
    public class ProfilePage : BasePage
    {
        //public IWebDriver Driver;
        public WebDriverWait Wait { get { return new WebDriverWait(Driver, TimeSpan.FromSeconds(60)); } }
        public UtilityHelper util = new UtilityHelper();

        public void GetTeamMemberPhoto(TeamMember tm)
        {
            By UWMButton = By.XPath("//body/div[@id='js-body-content']/div[@id='header']/div[1]/nav[1]/ul[1]/li[1]/a[1]");
            By allPeopleButton = By.XPath("//a[contains(text(),'All People')]");
            By search = By.XPath("//input[@id='txtSearchTerm']");
            By profileLink = By.XPath("//a[contains(text(),'" + tm.Name.Trim() + "')]");
            By image = By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/sk-profilepicture[1]");
            By profilePicture = By.XPath("//body/img[1]");
            By noProfile = By.XPath("//td[contains(text(),'Please enter a search term, or select all users fr')]");

            try
            {
                Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/");
                Wait.Until(d => d.FindElement(UWMButton).Displayed);
                Driver.FindElement(UWMButton).Click();
                Wait.Until(d => d.FindElement(allPeopleButton).Displayed);
                Driver.FindElement(allPeopleButton).Click();
                Wait.Until(d => d.FindElement(search).Displayed);
                Driver.FindElement(search).Click();
                Driver.FindElement(search).SendKeys(tm.Name);
                Driver.FindElement(search).SendKeys(Keys.Enter);

                //If Team Member's profile can not be found via search this changes their name to null so they can be removed from the list. 
                if (WaitForVisible(profileLink) || !WaitForVisible(noProfile))
                {
                    Driver.FindElement(profileLink).Click();
                    Wait.Until(d => d.FindElement(image).Displayed);
                    tm.PhotoPath = Driver.FindElement(image).GetAttribute("img");
                    Driver.Navigate().GoToUrl(@"https://uzone.unitedshore.com/user" + tm.PhotoPath);
                    IWebElement photo = Driver.FindElement(profilePicture);
                    tm.PhotoFilePath = util.TakeScreenshotOfElement(Driver, photo, tm.Name, tm.Location, tm.Id, true);
                }
                else
                {
                    tm.Name = null;
                }

            }
            catch (Exception)
            {
                Log.Information("Automation in GetTeamMemberPhoto() has failed");
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
