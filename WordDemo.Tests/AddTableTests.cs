using NPOI.XWPF.UserModel;
using Xunit;
using System.IO;

namespace WordDemo.Tests
{
    public class AddTableTests
    {
        [Fact]
        public void CreateTable_ShouldGenerateTableWithCorrectDimensions()
        {
            // Arrange
            var wordDoc = new XWPFDocument();

            // Act
            var table = wordDoc.CreateTable(3, 2);

            // Assert
            Assert.Equal(3, table.Rows.Count);
            Assert.Equal(2, table.Rows[0].GetTableCells().Count);
        }

        [Fact]
        public void HideBorders_ShouldRemoveAllBorders()
        {
            // Arrange
            var wordDoc = new XWPFDocument();
            var table = wordDoc.CreateTable(3, 2);

            // Act
            table.SetBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetInsideHBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");
            table.SetInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "WHITE");

            // Assert
            Assert.Equal(XWPFTable.XWPFBorderType.NONE, table.GetCTTbl().tblPr.tblBorders.bottom.border);
            Assert.Equal(XWPFTable.XWPFBorderType.NONE, table.GetCTTbl().tblPr.tblBorders.top.border);
            Assert.Equal(XWPFTable.XWPFBorderType.NONE, table.GetCTTbl().tblPr.tblBorders.left.border);
            Assert.Equal(XWPFTable.XWPFBorderType.NONE, table.GetCTTbl().tblPr.tblBorders.right.border);
        }

        [Fact]
        public void AddContentToCell_ShouldAddStyledText()
        {
            // Arrange
            var wordDoc = new XWPFDocument();
            var table = wordDoc.CreateTable(3, 2);
            var cell = table.GetRow(0).GetCell(0);

            // Act
            var paragraph = cell.AddParagraph();
            var run = paragraph.CreateRun();
            run.SetText("Some random text...");
            run.SetColor("FF5000");
            run.FontFamily = "Josefin Sans";
            run.FontSize = 14;
            run.IsBold = true;

            // Assert
            Assert.Equal("Some random text...", run.GetText(0));
            Assert.Equal("FF5000", run.GetColor());
            Assert.Equal("Josefin Sans", run.FontFamily);
            Assert.Equal(14, run.FontSize);
            Assert.True(run.IsBold);
        }
    }
}
