using System;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using Serilog;
using NUnit.Framework;
using USFS.Library.TestAutomation;
using USFS.Library.TestAutomation.Test;
using USFS.Library.TestAutomation.Util;
using UzonePageObject;

namespace Test
{
    [TestFixture]
    public class UzoneReport : BaseTest
    {
        private const string keyConnectionString = "ConnectionString";
        private const string keyUsername = "User";
        private const string keyPassword = "Password";
        private LoginPage loginPage;


        [OneTimeSetUp]
        public void OneTimeSetup()
        {
            
        }

        [SetUp]
        public void SetUp()
        {
            loginPage = new LoginPage();
            BasePage.StartBrowser(config[keyBrowserChoice], config[keyBrowserMode], config[keyResourceUsed]);
        }

        [Test]
        public void UzoneReportGeneration()
        {
            //Login
            LoginUzoneTest();
            //Connect DB and fetch records
            //Goto seatingmap and take screenshot
            //Create report in word format
        }

        public void LoginUzoneTest()
        {
            try
            {
                Log.Information("LoginUzoneTest has started");
                string pageTitle = loginPage.LoginUzone(config[keyUrl], config[keyUsername], config[keyPassword]);
                Log.Information("URL - " + config[keyUrl]);
                ReportingUtil.test.Info("URL - " + config[keyUrl]);
                StringAssert.AreEqualIgnoringCase("✔ YourTime", pageTitle, "Page title should contain '✔ YourTime'");
                Log.Information("LoginUzoneTest has passed");
                ReportingUtil.test.Pass("LoginUzoneTest has passed");
            }
            catch (Exception e)
            {
                Log.Error(e, "LoginUzoneTest Failed");
                throw new Exception("LoginUzoneTest Failed", e);
            }
        }
    }
}


    