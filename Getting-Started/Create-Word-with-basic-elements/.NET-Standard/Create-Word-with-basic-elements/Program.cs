using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Create_Word_with_basic_elements
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of WordDocument Instance (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Adds a new section into the Word document.
                IWSection section = document.AddSection();
                //Specifies the page margins.
                section.PageSetup.Margins.All = 50f;

                //Adds a new simple paragraph into the section.
                IWParagraph firstParagraph = section.AddParagraph();
                //Sets the paragraph's horizontal alignment as justify.
                firstParagraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Justify;
                //Adds a text range into the paragraph.
                IWTextRange firstTextRange = firstParagraph.AppendText("AdventureWorks Cycles,");
                //sets the font formatting of the text range.
                firstTextRange.CharacterFormat.Bold = true;
                firstTextRange.CharacterFormat.FontName = "Calibri";
                firstTextRange.CharacterFormat.FontSize = 14;
                //Adds another text range into the paragraph.
                IWTextRange secondTextRange = firstParagraph.AppendText(" the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //sets the font formatting of the text range.
                secondTextRange.CharacterFormat.FontName = "Calibri";
                secondTextRange.CharacterFormat.FontSize = 11;

                //Adds another paragraph and aligns it as center.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                //Sets after spacing for paragraph.
                paragraph.ParagraphFormat.AfterSpacing = 8;
                //Adds a picture into the paragraph.
                FileStream image1 = new FileStream(Path.GetFullPath(@"../../../Data/DummyProfilePicture.jpg"), FileMode.Open, FileAccess.Read);
                IWPicture picture = paragraph.AppendPicture(image1);
                //Specify the size of the picture.
                picture.Height = 100;
                picture.Width = 100;

                //Adds a table into the Word document.
                IWTable table = section.AddTable();
                //Creates the specified number of rows and columns.
                table.ResetCells(2, 2);
                //Accesses the instance of the cell (first row, first cell).
                WTableCell firstCell = table.Rows[0].Cells[0];
                //Specifies the width of the cell.
                firstCell.Width = 150;
                //Adds a paragraph into the cell; a cell must have atleast 1 paragraph.
                paragraph = firstCell.AddParagraph();
                IWTextRange textRange = paragraph.AppendText("Profile picture");
                textRange.CharacterFormat.Bold = true;
                //Accesses the instance of cell (first row, second cell).
                WTableCell secondCell = table.Rows[0].Cells[1];
                secondCell.Width = 330;
                paragraph = secondCell.AddParagraph();
                textRange = paragraph.AppendText("Description");
                textRange.CharacterFormat.Bold = true;
                firstCell = table.Rows[1].Cells[0];
                firstCell.Width = 150;
                paragraph = firstCell.AddParagraph();
                //Sets after spacing for paragraph.
                paragraph.ParagraphFormat.AfterSpacing = 6;
                FileStream image2 = new FileStream(Path.GetFullPath(@"../../../Data/DummyProfile-Picture.jpg"), FileMode.Open, FileAccess.Read);
                IWPicture profilePicture = paragraph.AppendPicture(image2);
                profilePicture.Height = 100;
                profilePicture.Width = 100;
                secondCell = table.Rows[1].Cells[1];
                secondCell.Width = 330;
                paragraph = secondCell.AddParagraph();
                textRange = paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

                //Writes default numbered list. 
                paragraph = section.AddParagraph();
                //Sets before spacing for paragraph.
                paragraph.ParagraphFormat.BeforeSpacing = 6;
                paragraph.AppendText("Level 0");
                //Applies the default numbered list formats.
                paragraph.ListFormat.ApplyDefNumberedStyle();
                //Applies list formatting.
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.LeftIndent = 36;
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.FirstLineIndent = -18;
                paragraph.ListFormat.CurrentListLevel.NumberAlignment = ListNumberAlignment.Left;
                paragraph = section.AddParagraph();
                paragraph.AppendText("Level 1");
                //Specifies the list format to continue from last list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increments the list level.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Applies list formatting.
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.LeftIndent = 72;
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.FirstLineIndent = -18;
                paragraph.ListFormat.CurrentListLevel.NumberAlignment = ListNumberAlignment.Left;
                paragraph = section.AddParagraph();
                paragraph.AppendText("Level 0");
                paragraph.ListFormat.ContinueListNumbering();
                //Decrements the list level.
                paragraph.ListFormat.DecreaseIndentLevel();
                //Applies list formatting.
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.LeftIndent = 36;
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.FirstLineIndent = -18;
                paragraph.ListFormat.CurrentListLevel.NumberAlignment = ListNumberAlignment.Left;
                section.AddParagraph();
                //Writes default bulleted list. 
                paragraph = section.AddParagraph();
                paragraph.AppendText("Level 0");
                //Applies the default bulleted list formats.
                paragraph.ListFormat.ApplyDefBulletStyle();
                //Applies list formatting.
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.LeftIndent = 36;
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.FirstLineIndent = -18;
                paragraph.ListFormat.CurrentListLevel.NumberAlignment = ListNumberAlignment.Left;
                paragraph = section.AddParagraph();
                paragraph.AppendText("Level 1");
                //Specifies the list format to continue from last list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increments the list level.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Applies list formatting.
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.LeftIndent = 72;
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.FirstLineIndent = -18;
                paragraph.ListFormat.CurrentListLevel.NumberAlignment = ListNumberAlignment.Left;
                paragraph = section.AddParagraph();
                paragraph.AppendText("Level 0");
                //Specifies the list format to continue from last list.
                paragraph.ListFormat.ContinueListNumbering();
                //Decrements the list level.
                paragraph.ListFormat.DecreaseIndentLevel();
                //Applies list formatting.
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.LeftIndent = 36;
                paragraph.ListFormat.CurrentListLevel.ParagraphFormat.FirstLineIndent = -18;
                paragraph.ListFormat.CurrentListLevel.NumberAlignment = ListNumberAlignment.Left;
                section.AddParagraph();

                //Creates file stream.
                using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(fileStream, FormatType.Docx);
                }
            }
        }
    }
}
