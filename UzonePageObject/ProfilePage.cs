﻿using OpenQA.Selenium;
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

            try
            {
                Log.Information("GetTeamMemberPhoto() has started.");
                Driver.Navigate().GoToUrl("https://uzone.unitedshore.com/user/index.html#/all");
                BrowserUtils.WaitForDisplayed(search, 30);
                Driver.FindElement(search).Click();
                Driver.FindElement(search).Clear();
                Driver.FindElement(search).SendKeys(tm.Name);
                Driver.FindElement(search).SendKeys(Keys.Enter);
                Thread.Sleep(1000);
                Log.Information("Waiting for profile link...");

                if (BrowserUtils.WaitForDisplayed(individualLink, 30) || !BrowserUtils.WaitForDisplayed(noProfile, 10))
                {
                    Log.Information("Team member profile link found - " + tm.Name.ToString());                    
                    Driver.FindElement(profileLink).Click();
                    BrowserUtils.WaitForDisplayed(image, 15);
                    tm.PhotoPath = Driver.FindElement(image).GetAttribute("img");
                    Driver.Navigate().GoToUrl(@"https://uzone.unitedshore.com/user" + tm.PhotoPath);
                    IWebElement photo = Driver.FindElement(profilePicture);
                    tm.PhotoFilePath = util.TakeScreenshotOfElement(Driver, photo, tm.Name, tm.Event, tm.Id, true);
                    Log.Information("Screenshot of profile pic element taken successfully - \nProfile Pic Location: " + tm.PhotoFilePath);
                }
                else
                {
                    //If Team Member's profile can not be found via search this changes their name to null. 
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