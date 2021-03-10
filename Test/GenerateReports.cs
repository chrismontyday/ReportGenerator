using System.Collections.Generic;
using UzonePageObject;
using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Drawing;
using Spire.Doc.Fields;
using Serilog;

namespace Test
{
    class GenerateReports
    {
        UtilityHelper auto = new UtilityHelper();

        public void CreateDocument(List<TeamMember> list, List<string> skippedProfs)
        {           
            Color blue = Color.FromArgb(0, 96, 154);
            Color green = Color.FromArgb(166, 195, 114);
            string generic = auto.ReturnPathFolder(3, "Testdata") + "generic.jpg";
            string filePath = @"\\dfs\ns1\BYOTL\Sandbox\Reports\" + 
                DateTime.Now.ToString("MMMM", System.Globalization.CultureInfo.InvariantCulture) + 
                "_Birthdays_Anniversary_" + 
                DateTime.Now.ToString("HH_mm_ss") + ".docx";
            string emailSubject = "Uzone Birthday/Anniversary Report Created: " + DateTime.Now.ToString("yyyy-MM-dd");
            string emailBody = "<a href=\"" + filePath + "\">Click Here for Report.</a>";            
            string fontName = "Roboto";

            try
            {
                Log.Information("CreateDocument() has started...");
                Document doc = new Document();
                Section one = doc.AddSection();
                Paragraph p1 = doc.Sections[0].AddParagraph();
                TextRange text1 = p1.AppendText(UtilityHelper.ConvertListToString(skippedProfs));
                Break pageBreak = new Break(doc, BreakType.PageBreak);

                foreach (TeamMember tm in list)
                {
                    string teamLead = tm.Event + "\nTeam Lead: " + tm.Name + "\n";
                    string subTeam = "\nSub Team: " + tm.SubTeamName +
                        "\nTeam: " + tm.TeamName + "\nFloor: " + tm.Floor;

                    if(tm.Skipped==false && tm.Name != null && tm.MapFilePath != null && tm.PhotoFilePath != null)
                    {
                        Log.Information(tm.Name + " is being added to report. Id: " + tm.Id);
                        Section s = doc.AddSection();

                        Table table = s.AddTable(true);
                        table.ResetCells(1, 2);
                        table.TableFormat.Borders.Color = green;
                        table.TableFormat.Borders.Vertical.Color = green;
                        Log.Information(tm.Id + " - Table created");

                        //Adds to cell 1
                        TextRange rangeOne = table[0, 0].AddParagraph().AppendText(teamLead);
                        rangeOne.CharacterFormat.FontName = fontName;
                        rangeOne.CharacterFormat.TextColor = blue;

                        DocPicture TMPhoto = table.Rows[0].Cells[0].Paragraphs[0].AppendPicture(auto.GetImage(tm.PhotoFilePath));
                        TMPhoto.Width = 115;
                        TMPhoto.Height = 115;
                        Log.Information(tm.Id + " - Added photo to table.");

                        TextRange rangeTwo = table[0, 0].AddParagraph().AppendText(subTeam);
                        rangeTwo.CharacterFormat.FontName = fontName;
                        rangeTwo.CharacterFormat.TextColor = blue;

                        //Adds to cell 2
                        DocPicture TMPhoto2 = table[0, 1].AddParagraph().AppendPicture(auto.GetImage(tm.MapFilePath));
                        TMPhoto2.Width = 350;
                        TMPhoto2.Height = 230;

                        DocPicture gen = table[0, 1].AddParagraph().AppendPicture(auto.GetImage(generic));
                        gen.Width = 350;
                        gen.Height = 153;

                        Break pageBreak2 = new Break(doc, BreakType.PageBreak);
                        Log.Information(tm.Name + " has been added to report.");
                    }
                }

                //Save report
                doc.SaveToFile(filePath, FileFormat.Docx);
                Log.Information("SUCCESS!!!! - Report created successfully.");              

                //Empties folders where screenshots & excel sheet were kept. 
                auto.ClearFolder(auto.ReturnPathFolder(3, "TestOutput\\Screenshots"));
                auto.ClearFolder(auto.ReturnPathFolder(3, "Exceldata"));
                Log.Information("Screenshots Folder has been emptied");

                //Sends report in email.
                //Email.SendEMail(emailSubject, emailBody, "CDAY@UWM.COM", "username@email.com", filePath);
                Log.Information("Email sent successfully!");

            }
            catch (Exception e)
            {
                Log.Information("Failed to create Document");
                throw new Exception("Failed in CreateDocument() : ", e);
            }
        }
    }
}