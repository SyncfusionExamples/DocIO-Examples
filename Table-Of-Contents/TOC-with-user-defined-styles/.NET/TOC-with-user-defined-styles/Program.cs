using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace TOC_with_user_defined_styles
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Creates a new custom styles.
                Style style = (WParagraphStyle)document.AddParagraphStyle("MyStyle");
                style.CharacterFormat.Bold = true;
                style.CharacterFormat.FontName = "Verdana";
                style.CharacterFormat.FontSize = 25;
                //Adds the section into the Word document.
                IWSection section = document.AddSection();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Adds the paragraph into the created section.
                IWParagraph paragraph = section.AddParagraph();
                //Appends the TOC field with LowerHeadingLevel and UpperHeadingLevel to determines the TOC entries.
                TableOfContent tableOfContents = paragraph.AppendTOC(1, 3);
                tableOfContents.UseHeadingStyles = false;
                //Sets the TOC level style based on the created TOC.
                tableOfContents.SetTOCLevelStyle(2, "MyStyle");
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section.
                paragraph = section.AddParagraph();
                //Adds the text for the headings.
                paragraph.AppendText("First Chapter");
                //Sets the built-in heading style.
                paragraph.ApplyStyle("MyStyle");
                //Adds the text into the paragraph.
                section.AddParagraph().AppendText(paraText);
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section.
                paragraph = section.AddParagraph();
                //Adds the text for the headings.
                paragraph.AppendText("Second Chapter");
                //Sets the built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading1);
                //Adds the text to the paragraph.
                section.AddParagraph().AppendText(paraText);
                //Adds the section into Word document.
                section = document.AddSection();
                //Adds a paragraph to a created section.
                paragraph = section.AddParagraph();
                //Adds the text for the headings.
                paragraph.AppendText("Third Chapter");
                //Sets the built-in heading style.
                paragraph.ApplyStyle("MyStyle");
                //Adds the text to the paragraph.
                section.AddParagraph().AppendText(paraText);
                //Updates the table of contents.
                document.UpdateTableOfContents();
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
