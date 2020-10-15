using OpenQA.Selenium;
using USFS.Library.TestAutomation;
using USFS.Library.TestAutomation.Util;

namespace UzonePageObject
{
    public class LoginPage : BasePage
    {
        public By UserNameField => By.XPath("//*[@id='UserName']");
        public By PasswordField => By.XPath("//input[@name='Password']");
        public By LoginButton => By.XPath("//INPUT[@id='login-btn']");

        public string LoginUzone(string url, string userName, string password)
        {
            Driver.Navigate().GoToUrl(url);
            Driver.FindElement(UserNameField).SendKeys(userName);
            Driver.FindElement(PasswordField).SendKeys(password);
            Driver.FindElement(LoginButton).Click();
            return Driver.Title;
        }
    }

}