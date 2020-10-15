using System;
    using System.Data;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Support.UI;
    using Serilog;
    using NUnit.Framework;
    using USFS.Library.TestAutomation;
    using USFS.Library.TestAutomation.Test;
    using USFS.Library.TestAutomation.Util;
//using UzonePageObject;

namespace Test
{
    [TestFixture]
    class UzoneReport : BaseTest
    {
    }


    public class YourtimeIntegrationTest : BaseTest
    {
        private const string keyConnectionString = "ConnectionString";
        private const string keyUsername = "User";
        private const string keyPassword = "Password";
        private const string keyPasswordLeader = "PasswordLeader";
        private const string keyPasswordLeaderSenior = "PasswordLeaderSenior";

        //private LandingPage landingPage;
        [OneTimeSetUp]
        public void OneTimeSetup()
        {

        }

        [SetUp]
        public void SetUp()
        {
            //landingPage = new LandingPage();


            BasePage.StartBrowser(config[keyBrowserChoice], config[keyBrowserMode], config[keyResourceUsed]);
        }

        [Test]
        public void YourtimeRegression()
        {

        }
    }
}