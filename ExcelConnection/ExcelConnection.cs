using System.Collections.Generic;
using System.IO;
using USFS.Library.TestAutomation.Util;
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

        List<TeamMember> list;

        public ExcelConnection()
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(FileHandler.GenerateDynamicFilePath(Path.Combine(@"Testdata\", @"UZone SeatingMap_BirthdayListApril_02_2020 - 7202020.xls")));
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets["UZone - Birthdays - SeatingMap"];
            xlRange = xlWorksheet.UsedRange;
        }

        public List<TeamMember> GetTeamMembers()
        {
            list = new List<TeamMember>();

            for (int row = 2; row < xlRange.Rows.Count + 1; row++)
            {
                if (xlRange.Cells[row, 1] != null)
                {
                    TeamMember teamMember = new TeamMember()
                    {
                        Name = (xlRange.Cells[row, 1] as Excel.Range).Value2.ToString(),
                        Email = (xlRange.Cells[row, 2] as Excel.Range).Value2.ToString(),
                        Location = (xlRange.Cells[row, 3] as Excel.Range).Value2.ToString(),
                        TeamName = (xlRange.Cells[row, 4] as Excel.Range).Value2.ToString(),
                        SubTeamName = (xlRange.Cells[row, 5] as Excel.Range).Value2.ToString(),
                        Id = row
                    };
                    list.Add(teamMember);
                }
                else
                {
                    row += xlRange.Columns.Count;
                }
            }

            xlWorkbook.Close(false);
            xlApp.Quit();
            return list;
        }
    }
}
