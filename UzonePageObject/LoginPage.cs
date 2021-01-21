using OpenQA.Selenium;
using System;
using System.Threading;
using USFS.Library.TestAutomation;
using USFS.Library.TestAutomation.Util;

namespace UzonePageObject
{
    public class LoginPage : BasePage
    {
        public By UserNameField => By.XPath("//*[@id='userNameInput']");
        public By PasswordField => By.XPath("//*[@id='passwordInput']");
        public By LoginButton => By.XPath("//*[@id='submitButton']");

        public string LoginUzone(string url, string userName, string password)
        {
            try
            {
                Driver.Navigate().GoToUrl(url);
                BrowserUtils.WaitForDisplayed(LoginButton,90);
                Driver.FindElement(UserNameField).SendKeys(userName);
                Driver.FindElement(PasswordField).SendKeys(password);
                Driver.FindElement(LoginButton).Click();
                Thread.Sleep(200);
                return Driver.Title;
            }
            catch(Exception e)
            {
                String a = e.Message.ToString();
                return Driver.Title;

            }
        }
    }
}