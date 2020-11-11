
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

namespace Test
{
    class GenerateReports
    {
        UtilityHelper auto = new UtilityHelper();


        public void CreateDocument(List<TeamMember> list)
        {
            string generic = @"C:\Users\cday\Pictures\bdayanni.jpg";
            string filePath = auto.ReturnPathFolder(3, "Reports\\") + "test.docx";

            Document doc = new Document();
            Log.Information("Created Document successfully.");
            Section one = doc.AddSection();
            Paragraph p = doc.Sections[0].AddParagraph();
            TextRange text1 = p.AppendText("Birthday & Anniversary Seating Map Report " + DateTime.Now.Date);
            Break pageBreak = new Break(doc, BreakType.PageBreak);

            foreach (TeamMember tm in list)
            {
                Section s = doc.AddSection();
                Table table = s.AddTable(true);
                table.ResetCells(1, 2);
                               
                DocPicture TMPhoto = table.Rows[1].Cells[2].Paragraphs[0].AppendPicture(GetImage(tm.PhotoFilePath));
                TMPhoto.Width = 750;
                TMPhoto.Height = 468;

                DocPicture TMMap = table.Rows[2].Cells[2].Paragraphs[0].AppendPicture(GetImage(tm.MapFilePath));
                TMMap.Width = 750;
                TMMap.Height = 468;

                DocPicture gen = table.Rows[2].Cells[2].Paragraphs[0].AppendPicture(GetImage(generic));
                gen.Width = 750;
                gen.Height = 468;

                Break pageBreak2 = new Break(doc, BreakType.PageBreak);
            }


            doc.SaveToFile(filePath, FileFormat.Docx);


        }

        public byte[] GetImage(string filePath)
        {
                Image img = Image.FromFile(filePath);
                byte[] arr;
                using (MemoryStream ms = new MemoryStream())
                {
                    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    arr = ms.ToArray();
                }

            return arr;
        }





























































































        //public void CreateDocument(List<TeamMember> list)
        //{
        //    string font = "Arial";
        //    string testPic = @"C:\Code\UzoneSelenium\QE-X-Basic-Selenium\QE-X-Basic-Selenium\TestOutput\2-aria ying-2020-11-05 13-05-35-north building 1st floor.jpg";
        //    string testMap = @"C:\Code\UzoneSelenium\QE-X-Basic-Selenium\QE-X-Basic-Selenium\TestOutput\5-seating_map-2020-11-05 13-07-58-north building 1st floor.jpg";
        //    string generic = @"C:\Users\cday\Pictures\bdayanni.jpg";
        //    string filePath = auto.ReturnPathFolder(3, "Reports\\") + "test.docx";
        //    int finalWidth = 1000;
        //    int finalHeight = 1000;

        //    using (WordprocessingDocument doc =
        //    WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        //    {
        //        MainDocumentPart mainPart = doc.AddMainDocumentPart();
        //        mainPart.Document = new Document();

        //        foreach (TeamMember tm in list)
        //        {
        //            Body body = mainPart.Document.AppendChild(new Body());
        //            Run rOne = body.AppendChild(new Run());

        //            RunProperties rPr = new RunProperties(new RunFonts() { Ascii = font });

        //            Run r = doc.MainDocumentPart.Document.Descendants<Run>().First();

        //            r.PrependChild<RunProperties>(rPr);

        //            //Creates table then sets properties
        //            Table table = body.AppendChild(new Table());
        //            TableProperties tblProp = new TableProperties(
        //            new TableBorders(
        //            new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 9, Color = "006199" },
        //            new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 9, Color = "006199" },
        //            new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 9, Color = "006199" },
        //            new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 9, Color = "006199" },
        //            new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 9, Color = "006199" },
        //            new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.BasicThinLines), Size = 9, Color = "006199" }
        //            )
        //            );
        //            table.AppendChild<TableProperties>(tblProp);

        //            //Images
        //            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
        //            using (FileStream stream = new FileStream(generic, FileMode.Open))
        //            {
        //                imagePart.FeedData(stream);
        //            }

        //            ImagePart teamMember = mainPart.AddImagePart(ImagePartType.Jpeg);
        //            using (FileStream stream = new FileStream(tm.PhotoFilePath, FileMode.Open))
        //            {
        //                teamMember.FeedData(stream);
        //            }
        //            ImagePart seatingMap = mainPart.AddImagePart(ImagePartType.Jpeg);
        //            using (FileStream stream = new FileStream(tm.MapFilePath, FileMode.Open))
        //            {
        //                seatingMap.FeedData(stream);
        //            }



        //            TableRow tr = new TableRow();
        //            TableCell tc1 = new TableCell();
        //            tc1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3200" }));
        //            tc1.Append(new Paragraph(new Run(new Text("Team Lead: " + tm.Name))));
        //            AddImageToCell(tc1, mainPart.GetIdOfPart(teamMember), tm.PhotoFilePath);
        //            tc1.Append(new Paragraph(new Run(new Text("SubTeam: " + tm.SubTeamName))));
        //            tc1.Append(new Paragraph(new Run(new Text("Team: " + tm.TeamName))));
        //            tr.Append(tc1);



        //            TableCell tc2 = new TableCell();
        //            tc2.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "6400" }));
        //            tc2.Append(new Paragraph(new Run(new Text("Seating Map"))));
        //            AddImageToCell(tc2, mainPart.GetIdOfPart(seatingMap), tm.MapFilePath);
        //            AddImageToCell(tc2, mainPart.GetIdOfPart(imagePart), generic);
        //            tr.Append(tc2);
        //            table.Append(tr);


        //            //SdtElement block = doc.MainDocumentPart.Document.Body.Descendants<SdtElement>()
        //            //    .FirstOrDefault(sdt => sdt.SdtProperties.GetFirstChild<Tag>()?.Val == mainPart.GetIdOfPart(seatingMap));

        //            //Drawing sdtImage = block.Descendants<Drawing>().First();
        //            //double sdtWidth = sdtImage.Inline.Extent.Cx;
        //            //double sdtHeight = sdtImage.Inline.Extent.Cy;
        //            //double sdtRatio = sdtWidth / sdtHeight;


        //            ////Resize picture placeholder
        //            //sdtImage.Inline.Extent.Cx = finalWidth;
        //            //sdtImage.Inline.Extent.Cy = finalHeight;

        //            ////Change width/height of picture shapeproperties Transform
        //            ////This will override above height/width until you manually drag image for example
        //            //sdtImage.Inline.Graphic.GraphicData
        //            //    .GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
        //            //    .ShapeProperties.Transform2D.Extents.Cx = finalWidth;
        //            //sdtImage.Inline.Graphic.GraphicData
        //            //    .GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
        //            //    .ShapeProperties.Transform2D.Extents.Cy = finalHeight;

        //            //Creates page break
        //            Paragraph para = table.AppendChild(new Paragraph(new Run((new Break() { Type = BreakValues.Page }))));

        //        }
        //    }
        //}

        ////This should set the font to Arial but instead crashes program
        //public static void SetRunFont(string fileName)
        //{
        //    string Rgb = "FFA500";
        //    // Open a Wordprocessing document for editing.
        //    using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, true))
        //    {
        //        // Set the font to Arial to the first Run.
        //        // Use an object initializer for RunProperties and rPr.
        //        DocumentFormat.OpenXml.Wordprocessing.Color color = new DocumentFormat.OpenXml.Wordprocessing.Color();
        //        color.Val = Rgb;
        //        RunProperties rPr = new RunProperties(
        //            new RunFonts()
        //            {
        //                Ascii = "Arial",

        //            });
        //        rPr.Append(color);

        //        Run r = package.MainDocumentPart.Document.Descendants<Run>().First();
        //        r.PrependChild<RunProperties>(rPr);

        //        // Save changes to the MainDocumentPart part.
        //        package.MainDocumentPart.Document.Save();
        //    }
        //}


        //private static void AddImageToCell(TableCell cell, string relationshipId, string strImageFile)
        //{

        //    long iImageWidth = 0, iImageHeight = 0, iOffset = 10;
        //    using (Bitmap bitmap = new Bitmap(strImageFile))
        //    {
        //        iImageHeight = bitmap.Height;
        //        iImageWidth = bitmap.Width;
        //    }



        //    var element =
        //    new Drawing(
        //    new DW.Inline(
        //    new DW.Extent() { Cx = iImageWidth, Cy = iImageHeight },
        //    new DW.EffectExtent()
        //    {
        //        LeftEdge = 0L,
        //        TopEdge = 0L,
        //        RightEdge = 0L,
        //        BottomEdge = 0L
        //    },
        //    new DW.DocProperties()
        //    {
        //        Id = (UInt32Value)1U,
        //        Name = "Picture 1"
        //    },
        //    new DW.NonVisualGraphicFrameDrawingProperties(
        //    new A.GraphicFrameLocks() { NoChangeAspect = true }),
        //    new A.Graphic(
        //    new A.GraphicData(
        //    new PIC.Picture(
        //    new PIC.NonVisualPictureProperties(
        //    new PIC.NonVisualDrawingProperties()
        //    {
        //        Id = (UInt32Value)0U,
        //        Name = "New Bitmap Image.jpg"
        //    },
        //    new PIC.NonVisualPictureDrawingProperties()),
        //    new PIC.BlipFill(
        //    new A.Blip(
        //    new A.BlipExtensionList(
        //    new A.BlipExtension()
        //    {
        //        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
        //    })
        //    )
        //    {
        //        Embed = relationshipId,
        //        CompressionState =
        //    A.BlipCompressionValues.Print
        //    },
        //    new A.Stretch(
        //    new A.FillRectangle())),
        //    new PIC.ShapeProperties(
        //    new A.Transform2D(
        //    new A.Offset() { X = 0L, Y = 0L },
        //    new A.Extents() { Cx = iImageWidth, Cy = iImageHeight }
        //    ),
        //    new A.PresetGeometry(
        //    new A.AdjustValueList()
        //    )
        //    { Preset = A.ShapeTypeValues.Rectangle }))
        //    )
        //    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
        //    )
        //    )
        //    //{
        //    //    DistanceFromTop = (UInt32Value)0U,
        //    //    DistanceFromBottom = (UInt32Value)0U,
        //    //    DistanceFromLeft = (UInt32Value)0U,
        //    //    DistanceFromRight = (UInt32Value)0U
        //    //}
        //    );

        //    cell.Append(new Paragraph(new Run(element)));
        //}

    }
}