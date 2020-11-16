using NUnit.Framework;
using Serilog;
using System;
using System.Collections.Generic;
using USFS.Library.TestAutomation;
using USFS.Library.TestAutomation.Test;
using USFS.Library.TestAutomation.Util;
using UzonePageObject;
using ExcelConnect;
using System.Threading;

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
        GenerateReports doc = new GenerateReports();
        UtilityHelper auto = new UtilityHelper();
        ExcelConnection excel;
        List<TeamMember> list;

        [SetUp]
        public void SetUp()
        {
            excel = new ExcelConnection();
            list = excel.GetTeamMembers();

            //This will kill all excel tasks running at the time it was called. 
            auto.KillExcel();
            Log.Information("taskkill - excel.exe ran. All open versions of Excel are now closed.");

            loginPage = new LoginPage();
            profilePage = new ProfilePage();
            seatingMap = new SeatingMapPage();
            BasePage.StartBrowser(config[keyBrowserChoice], config[keyBrowserMode], config[keyResourceUsed]);
        }

        [TearDown]
        public void TearDown()
        {
            BasePage.CloseBrowser();
        }

        [Test]
        public void UzoneReportGeneration()
        {
            //Login
            LoginUzoneTest();
            //Goto seatingmap and take screenshot
            GetPicsFromUzone(list);
            //Create report in word format
            doc.CreateDocument(list);
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
            try
            {
                Log.Information("GetPicsFromUzone() has started");
                foreach (TeamMember tm in list)
                {
                    try
                    {
                        Log.Information("Getting Profile Pic & Seating Map for - " + tm.Id);
                        if (tm.Name != "Leader" && tm.Name != null)
                        {
                            profilePage.GetTeamMemberPhoto(tm);

                            if (tm.Name != null)
                            {
                                seatingMap.GetTeamMemberSeatingMap(tm);
                            }
                        }
                        Log.Information("Success! - " + tm.Id);
                        ReportingUtil.test.Pass("GetPics() has passed");

                    }
                    catch
                    {
                        Log.Information("GetPics() failed on team member - Id: " + tm.Id);
                    }
                }
            }
            catch (Exception e)
            {
                Log.Error(e, "GetPicsFromUzone() has Failed");
                throw new Exception("GetPicsFromUzone() Failed", e);
            }
        }       

    }
}


