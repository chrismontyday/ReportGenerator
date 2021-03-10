using NUnit.Framework;
using Serilog;
using System;
using System.Collections.Generic;
using USFS.Library.TestAutomation;
using USFS.Library.TestAutomation.Test;
using USFS.Library.TestAutomation.Util;
using UzonePageObject;
using ExcelConnect;

namespace Test
{
    /// <summary>
    /// Uzone POM takes an excel sheet from /Exceldata > Reads the information and creates a List<TeamMembers> 
    /// which contains all of the information from the Excel Sheet > Using this data, the bot searches for 
    /// the Team Member's photo on Uzone.com and their seating map > then finally generates reports on all of
    /// for every team member and saves it to a file. 
    /// </summary>

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
        static List<List<TeamMember>> multiList;
        static List<TeamMember> list;

        [SetUp]
        public void SetUp()
        {
            Log.Information("Program Starting...");
            excel = new ExcelConnection();
            list = excel.GetTeamMembers();

            //This will kill all excel tasks running at the time it is called. 
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
            //Gets Team Member's Profile Pics & Seating Map 
            UzoneOrchestrator(list);
        }

        public void LoginUzoneTest()
        {
            try
            {
                Log.Information("LoginUzoneTest has started");
                string pageTitle = loginPage.LoginUzone(config[keyUrl], config[keyUsername], config[keyPassword]);
                Log.Information("URL - " + config[keyUrl]);
                ReportingUtil.test.Info("URL - " + config[keyUrl]);
                Log.Information("LoginUzoneTest has passed");
                ReportingUtil.test.Pass("LoginUzoneTest has passed");
            }
            catch (Exception e)
            {
                Log.Error(e, "LoginUzoneTest Failed");
                throw new Exception("LoginUzoneTest Failed", e);
            }
        }

        public void GetPicsFromUzone(List<TeamMember> GetPicsList)
        {
            List<string> skippedNotes = new List<string>();

            foreach (TeamMember tm in GetPicsList)
            {
                Log.Information(tm.Id + " - " + tm.Name + " - Started automation");
                try
                {

                    if (tm.Name != "Leader" && tm.Name != null && tm.Skipped == false)
                    {
                        profilePage.GetTeamMemberPhoto(tm);

                        if (tm.Name != null || tm.Skipped == false)
                        {
                            seatingMap.GetTeamMemberSeatingMap(tm); 
                        }
                    }

                    if (tm.Skipped == true)
                    {
                        skippedNotes.Add(tm.Name + " was skipped. Message:" + tm.SkippedNote);
                        Log.Information(tm.Id + " - Failed...");
                    }
                    else
                    {
                        Log.Information(tm.Id + " - Success!");
                        ReportingUtil.test.Pass("GetPics() has passed");
                    }

                }
                catch
                {
                    Log.Information("GetPicsFromUzone() failed on team member - Id: " + tm.Id + " - " + tm.SkippedNote);
                }
            }

            doc.CreateDocument(GetPicsList, skippedNotes);            
            skippedNotes.Clear();
        }



        /// <summary>
        /// This Method was added to reduce the amount of memory that the bot was using up. It was breaking after about 300 people. 
        /// </summary>
        public void UzoneOrchestrator(List<TeamMember> list)
        {
            List<TeamMember> chunkOfList;
            // 175 was chosen because it is known than the bot can handle that amount. 
            int num = 175;

            if (list.Count > num)
            {
                int timesToSplit = list.Count / num;

                for (int i = 0; i < timesToSplit; i++)
                {
                    chunkOfList = new List<TeamMember>();

                    for (int j = 0; j < num; j++)
                    {
                        chunkOfList.Add(list[j]);
                    }       
                    
                    GetPicsFromUzone(chunkOfList);
                    chunkOfList.Clear();
                    list.RemoveRange(0, num);
                }

                GetPicsFromUzone(list);
            }
            else
            {
                GetPicsFromUzone(list);
            }
        }
    }
}


