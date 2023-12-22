using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.XWPF.Model;
using NPOI.XWPF.UserModel;
using SixLabors.ImageSharp;
using System.Reflection.Metadata;
using System.Xml.Linq;

namespace WordDemo
{
    public class Program
    {
        public const string fillerText1 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";
        public const string fillerText2 = "The quick brown fox jumped over a lazy dog!";

        static void Main(string[] args)
        {
            XWPFDocument wordDoc = new XWPFDocument();


            XWPFHeaderFooterPolicy headerFooterPolicy = wordDoc.GetHeaderFooterPolicy();
            if (headerFooterPolicy == null) headerFooterPolicy = wordDoc.CreateHeaderFooterPolicy();

            // create header and add a test header
            XWPFHeader header = headerFooterPolicy.CreateHeader(XWPFHeaderFooterPolicy.DEFAULT);
            XWPFParagraph paragraph1 = header.CreateParagraph();
            paragraph1.Alignment = (ParagraphAlignment.LEFT);
            XWPFRun run1 = paragraph1.CreateRun();
            run1.SetText("Header Test");

            // create footer
            XWPFFooter footer = headerFooterPolicy.CreateFooter(XWPFHeaderFooterPolicy.DEFAULT);

            // add power logo
            var paragraph3 = footer.CreateParagraph();
            paragraph3.Alignment = (ParagraphAlignment.RIGHT);

            // add copyright
            var paragraph2 = footer.CreateParagraph();
            paragraph2.Alignment = (ParagraphAlignment.CENTER);
            XWPFRun run2 = paragraph2.CreateRun();
            run2.SetText("© 2024 Devon Software & Services Pvt. Ltd.");

            // Add an image in the header
            const string imageName = "sampleImage.png";
            XWPFRun run = AddImage(paragraph1.CreateRun(), imageName, 50, 50);

            // Add Page Margins
            CT_SectPr sectPr = wordDoc.Document.body.sectPr;
            CT_PageMar pageMar = sectPr.AddPageMar();
            pageMar.left = 720; //720 Twentieths of an Inch Point (Twips) = 720/20 = 36 pt = 36/72 = 0.5"
            pageMar.right = 720;
            pageMar.top = 720; //1440 Twips = 1440/20 = 72 pt = 72/72 = 1"
            pageMar.bottom = 720;
            pageMar.header = 72; //45.4 pt * 20 = 908 = 45.4 pt header from top
            pageMar.footer = 72; //28.4 pt * 20 = 568 = 28.4 pt footer from bottom


            //Add a Table

            // Create a table
            XWPFTable exampleTable = wordDoc.CreateTable(3, 2);
            exampleTable.SetColumnWidth(0, 3500);
            exampleTable.SetColumnWidth(1, 7000);

            // Set Table Layout to fixed
            SetTableLayoutToFixed(exampleTable);

            // Hide table borders
            exampleTable = HideBorders(exampleTable);

            // Add contents to a cell
            XWPFTableCell cell1 = exampleTable.GetRow(0).GetCell(0);
            XWPFParagraph paragraph4 = cell1.AddParagraph();
            paragraph4.Alignment = ParagraphAlignment.CENTER;
            XWPFRun run4 = paragraph4.CreateRun();
            run4.SetText(" Some random text...");
            run4.SetColor("FF5000");
            run4.FontFamily = "Josefin Sans";
            run4.FontSize = 14;
            run4.IsBold = true;

            // Simple bullet example

            SimpleBulletExample(wordDoc, cell1);

            //Multilevel bullet example

            MultiLevelBulletExample(wordDoc);


            string outputFile = AppDomain.CurrentDomain.BaseDirectory + "TestWord" + ".docx";
            FileStream resumeOutputFile = new FileStream(outputFile, FileMode.Create);
            wordDoc.Write(resumeOutputFile);
        }

        public static XWPFRun AddImage(XWPFRun run, string imageName, int width, int height)
        {
            var widthEmus = (int)(width * 9525 * 0.75);
            var heightEmus = (int)(height * 9525 * 0.75);

            using (FileStream picData = new FileStream(imageName, FileMode.Open, FileAccess.Read))
            {
                run.AddPicture(picData, (int)NPOI.XWPF.UserModel.PictureType.PNG, imageName, widthEmus, heightEmus);
            }

            return run;
        }

        public static void SetTableLayoutToFixed(XWPFTable table)
        {
            CT_TblLayoutType tblLayout1 = table.GetCTTbl().tblPr.AddNewTblLayout();
            tblLayout1.type = ST_TblLayoutType.@fixed;
        }

        public static XWPFTable HideBorders(XWPFTable table)
        {
            // Vanish borders
            table.SetBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            return table;
        }


        public static void SimpleBulletExample(XWPFDocument doc, XWPFTableCell cell)
        {
            XWPFNumbering numbering = doc.CreateNumbering();
            string abstractNumId = numbering.AddAbstractNum();
            string numId = numbering.AddNum(abstractNumId);

            XWPFParagraph p0 = cell.AddParagraph();
            XWPFRun r0 = p0.CreateRun();
            r0.SetText(fillerText2);
            r0.FontFamily = "Josefin Sans";
            r0.FontSize = 10;
            p0.SetNumID(numId);

            XWPFParagraph p1 = cell.AddParagraph();
            XWPFRun r1 = p1.CreateRun();
            r1.SetText(fillerText2);
            r1.FontFamily = "Josefin Sans";
            r1.FontSize = 10;
            p1.SetNumID(numId);
        }

        public static void MultiLevelBulletExample(XWPFDocument doc)
        {
            XWPFNumbering numbering = doc.CreateNumbering();
            var abstractNumId = numbering.AddAbstractNum();
            var numId = numbering.AddNum(abstractNumId);
            doc.CreateParagraph();
            doc.CreateParagraph();

            var p1 = doc.CreateParagraph();
            var r1 = p1.CreateRun();
            r1.SetText("multi level bullet");
            r1.IsBold = true;
            r1.FontFamily = "Courier";
            r1.FontSize = 12;

            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first");
            p1.SetNumID(numId, "0");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first-first");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first-second");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("first-third");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second");
            p1.SetNumID(numId, "0");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-first");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-second");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-third");
            p1.SetNumID(numId, "1");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-third-first");
            p1.SetNumID(numId, "2");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("second-third-second");
            p1.SetNumID(numId, "2");
            p1 = doc.CreateParagraph();
            r1 = p1.CreateRun();
            r1.SetText("third");
            p1.SetNumID(numId, "0");
        }
    }
}
