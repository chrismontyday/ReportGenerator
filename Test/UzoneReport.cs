using NUnit.Framework;
using Serilog;
using System;
using System.Collections.Generic;
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
        private ProfilePage profilePage;
        private SeatingMapPage seatingMap;
        UtilityHelper auto = new UtilityHelper();
        ExcelConnection excel;
        List<TeamMember> list; 

        [OneTimeSetUp]
        public void OneTimeSetup()
        {

        }

        [SetUp]
        public void SetUp()
        {
            excel = new ExcelConnection();
            list = PopulateTeamMemberList();
            loginPage = new LoginPage();
            BasePage.StartBrowser(config[keyBrowserChoice], config[keyBrowserMode], config[keyResourceUsed]);
        }

        [Test]
        public void UzoneReportGeneration()
        {
            //Login
            LoginUzoneTest();
            //Goto seatingmap and take screenshot
            GetPicsFromUzone(list);
            //Create report in word format
        }

        public void LoginUzoneTest()
        {
            try
            {
                Log.Information("LoginUzoneTest has started");
                string pageTitle = loginPage.LoginUzone(config[keyUrl], config[keyUsername], config[keyPassword]);
                //string pageTitle = loginPage.GotoUzoneHomePage(config[keyUrl]);
                Log.Information("URL - " + config[keyUrl]);
                ReportingUtil.test.Info("URL - " + config[keyUrl]);
                //StringAssert.AreEqualIgnoringCase("✔ YourTime", pageTitle, "Page title should contain '✔ YourTime'");
                Log.Information("LoginUzoneTest has passed");
                ReportingUtil.test.Pass("LoginUzoneTest has passed");
            }
            catch (Exception e)
            {
                Log.Error(e, "LoginUzoneTest Failed");
                throw new Exception("LoginUzoneTest Failed", e);
            }
        }

        public void GetPicsFromUzone(List<TeamMember> list)
        {
            auto.SetUpWebDriver();

            foreach (TeamMember tm in list)
            {
                if (tm.Name != "Leader" && tm.Name != null)
                {
                    profilePage.GetTeamMemberPhoto(tm);

                    if (tm.Name != null)
                    {
                        seatingMap.GetTeamMemberSeatingMap(tm);
                    }
                }
            }
            auto.TeardownWebDriver();
        }

        public List<TeamMember> PopulateTeamMemberList()
        {
            list = excel.GetTeamMembers();
            excel.CloseWorkbook();
            return list;
        }
    }
}


