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
            string[] file =  Directory.GetFiles(auto.ReturnPathFolder(3, @"Exceldata"));
            return file[0].ToString();
        }

        public ExcelConnection(string sheetName = "UZone - Birthdays - SeatingMap")
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

            try
            {
                Log.Information("Reading excel sheet");
                for (int row = 2; row < xlRange.Rows.Count + 1; row++)
                {
                    if (xlRange.Cells[row, 1] != null)
                    {
                        TeamMember teamMember = new TeamMember()
                        {
                            Name = (xlRange.Cells[row, 1] as Excel.Range).Value2.ToString(),
                            Event = (xlRange.Cells[row, 2] as Excel.Range).Value2.ToString(),
                            Email = (xlRange.Cells[row, 3] as Excel.Range).Value2.ToString(),
                            TeamName = (xlRange.Cells[row, 4] as Excel.Range).Value2.ToString(),
                            SubTeamName = (xlRange.Cells[row, 5] as Excel.Range).Value2.ToString(),
                            Id = row
                        };
                        list.Add(teamMember);
                        Log.Information(teamMember.Name + " has been added to the list from excel sheet.");
                    }
                    else
                    {
                        row += xlRange.Columns.Count;
                    }
                }

                //These don't do anything... I have no idea why. 
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
