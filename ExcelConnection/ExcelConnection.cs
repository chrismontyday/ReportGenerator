using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using UzonePageObject;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelConnect
{
    public class ExcelConnection
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;
        UtilityHelper auto = new UtilityHelper();
        List<TeamMember> list;

        public string GetFile()
        {
            string[] file = Directory.GetFiles(auto.ReturnPathFolder(3, @"Exceldata"));
            return file[0].ToString();
        }

        public ExcelConnection(string sheetName = "Sheet1")
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(GetFile());
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[@sheetName];
            xlRange = xlWorksheet.UsedRange;
            Log.Information("Excel Application Opened");
        }

        public List<TeamMember> GetTeamMembers()
        {
            list = new List<TeamMember>();

            int nameColumn = 1;
            int emailColumn = 2;
            int teamColumn = 3;
            int subteamColumn = 4;

            for (int i = 1; i < xlRange.Columns.Count + 1; i++)
            {
                string columnName = (xlRange.Cells[1, i] as Excel.Range).Value2.ToString();

                if (columnName.ToLower().Contains("leader") || columnName.ToLower() == "name")
                {
                    nameColumn = i;
                }
                else if (columnName.ToLower().Contains("mail"))
                {
                    emailColumn = i;
                }
                else if (columnName.ToLower() == "team name" || columnName.ToLower() == "teamname")
                {
                    teamColumn = i;
                }
                else if (columnName.ToLower().Contains("sub"))
                {
                    subteamColumn = i;
                }
            }

            try
            {
                Log.Information("Reading excel sheet");
                for (int row = 2; row < xlRange.Rows.Count + 1; row++)
                {
                    if (xlRange.Cells[row, 1] != null)
                    {
                        TeamMember teamMember = new TeamMember()
                        {
                            Name = (xlRange.Cells[row, nameColumn] as Excel.Range).Value2.ToString(),
                            Email = (xlRange.Cells[row, emailColumn] as Excel.Range).Value2.ToString(),
                            TeamName = (xlRange.Cells[row, teamColumn] as Excel.Range).Value2.ToString(),
                            SubTeamName = (xlRange.Cells[row, subteamColumn] as Excel.Range).Value2.ToString(),
                            Id = row
                        };
                        list.Add(teamMember);
                        Log.Information(teamMember.Name + " has been red from excel sheet.");
                    }
                    else
                    {
                        row += xlRange.Columns.Count;
                    }
                }

                //These are supposed to close excel but they don't do anything... I have no idea why. 
                //There is a command prompt that runs to kill the processes in the UzoneReport Class immediately after this is called. 
                xlWorkbook.Close();
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
                Log.Information("Excel Application Closed.");
            }
            catch (Exception e)
            {
                //Why doesn't excel close? hmm...
                xlWorkbook.Close();
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
                Log.Information("Excel Application Closed.");
                Log.Error(e, "GetTeamMembers() has failed");
                throw new Exception("GetTeamMembers() Failed", e);
            }

            return list;
        }
    }
}
