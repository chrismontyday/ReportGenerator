using OpenQA.Selenium;
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
            Driver.Navigate().GoToUrl(url);
            Driver.FindElement(UserNameField).SendKeys(userName);
            Driver.FindElement(PasswordField).SendKeys(password);
            Driver.FindElement(LoginButton).Click();
            Thread.Sleep(200);
            return Driver.Title;
        }


        //Autologin with SSO
        public string GotoUzoneHomePage(string url)
        {
            Driver.Navigate().GoToUrl(url);
            Thread.Sleep(60);
            return Driver.Title;
        }
    }

}