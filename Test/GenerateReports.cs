
using System.Collections.Generic;
using System.Linq;
using UzonePageObject;
using Spire.Doc;
using Spire.Doc.Documents;

using System;

using System.Data;

using System.Drawing;


using Spire.Doc.Fields;
using Serilog;
using System.IO;
using OfficeOpenXml.Drawing;

namespace Test
{
    class GenerateReports
    {
        UtilityHelper auto = new UtilityHelper();

        public void CreateDocument(List<TeamMember> list)
        {
            string generic = auto.ReturnPathFolder(3, "Testdata") + "generic.jpg";
            string filePath = auto.ReturnPathFolder(3, "Reports") + list.Count.ToString() + "-SeatingMap-" + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".docx";

            try
            {
                Log.Information("CreateDocument() has started...");
                Document doc = new Document();
                Section one = doc.AddSection();
                Paragraph p1 = doc.Sections[0].AddParagraph();
                TextRange text1 = p1.AppendText("Birthday & Anniversary Seating Map Report " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss"));
                Break pageBreak = new Break(doc, BreakType.PageBreak);

                foreach (TeamMember tm in list)
                {
                    if (tm.Name != null)
                    {
                        Log.Information(tm.Name + " is being added to report.");
                        Section s = doc.AddSection();

                        Table table = s.AddTable(true);
                        table.ResetCells(1, 2);
                        table.TableFormat.Borders.Color = Color.FromArgb(0, 96, 154);
                        table.TableFormat.Borders.Vertical.Color = Color.FromArgb(0, 96, 154);

                        //Adds to cell 1
                        Log.Information(tm.Name + " photo has been added to report.");
                        TextRange rangeOne = table[0, 0].AddParagraph().AppendText(tm.Event + "\nTeam Lead: " + tm.Name + "\n");
                        rangeOne.CharacterFormat.FontName = "Calibri";
                        rangeOne.CharacterFormat.TextColor = Color.FromArgb(255, 165, 0);

                        DocPicture TMPhoto = table.Rows[0].Cells[0].Paragraphs[0].AppendPicture(auto.GetImage(tm.PhotoFilePath));
                        TMPhoto.Width = 115;
                        TMPhoto.Height = 115;

                        TextRange rangeTwo = table[0, 0].AddParagraph().AppendText("\nSub Team: " + tm.SubTeamName +
                        "\nTeam: " + tm.TeamName);
                        rangeTwo.CharacterFormat.FontName = "Calibri";
                        rangeTwo.CharacterFormat.TextColor = Color.FromArgb(255, 165, 0);

                        //Adds to cell 2
                        Log.Information(tm.Name + " seating map has been added to report.");
                        DocPicture TMPhoto2 = table[0, 1].AddParagraph().AppendPicture(auto.GetImage(tm.MapFilePath));
                        TMPhoto2.Width = 350;
                        TMPhoto2.Height = 230;

                        DocPicture gen = table[0, 1].AddParagraph().AppendPicture(auto.GetImage(generic));
                        gen.Width = 350;
                        gen.Height = 163;

                        Break pageBreak2 = new Break(doc, BreakType.PageBreak);
                        Log.Information(tm.Name + " index has been added to report.");
                    }
                }

                //Save report
                doc.SaveToFile(filePath, FileFormat.Docx);
                Log.Information("Created Document successfully.");

                //Empties out folder where screenshots were kept. 
                auto.ClearFolder(auto.ReturnPathFolder(3, "TestOutput\\Screenshots"));
                Log.Information("Screenshots Folder has been emptied");
            }
            catch (Exception e)
            {
                Log.Error(e, "CreateDocument() has Failed");
                throw new Exception("CreateDocument() Failed", e);
            }
        }
    }
}