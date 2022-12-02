using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Apply_switches_for_TOC
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section into the Word document.
                IWSection section = document.AddSection();
                //Add a paragraph into the created section.
                IWParagraph paragraph = section.AddParagraph();
                // Append table of content to the end of the paragraph with the specified lower and upper heading levels.
                TableOfContent tableOfContent = paragraph.AppendTOC(2, 5);
                //Set starting heading level of the table of contents.
                tableOfContent.LowerHeadingLevel = 2;
                //Set ending heading level of the table of contents.
                tableOfContent.UpperHeadingLevel = 5;
                //Use default heading styles
                tableOfContent.UseHeadingStyles = true;
                //Show page numbers in table of content.
                tableOfContent.IncludePageNumbers = true;
                //Set page numbers to right alignment.
                tableOfContent.RightAlignPageNumbers = false;
                //Set hyperlinks for the TOC levels.
                tableOfContent.UseHyperlinks = false;
                //Add a paragraph into the section.
                paragraph = section.AddParagraph();
                //Append text.
                paragraph.AppendText("First ");
                //Append line break.
                paragraph.AppendBreak(BreakType.LineBreak);
                paragraph.AppendText("Chapter");
                //Include new line to preserve newline characters TableOfContent.
                tableOfContent.IncludeNewLineCharacters = true;
                //Set a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading2);
                //Add a section into the Word document.
                section = document.AddSection();
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Append text.
                paragraph.AppendText("Second ");
                paragraph.AppendText("Chapter");
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading1);
                //Add a section into the Word document.
                section = document.AddSection();
                paragraph = section.AddParagraph();
                paragraph.AppendText("Third ");
                paragraph.AppendText("Chapter");
                //Set a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading2);
                section = document.AddSection();
                paragraph = section.AddParagraph();
                paragraph.AppendText("Fourth ");
                paragraph.AppendText("Chapter");
                //Set a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading3);
                //Add the text into the new paragraph of the section.
                section.AddParagraph().AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Set outline level for paragraph.
                paragraph.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
                //Append text.
                paragraph.AppendText("Outline Level Paragraph");
                //Set the outline levels.
                tableOfContent.UseOutlineLevels = true;
                //Add a section into the Word document.
                section = document.AddSection();
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Append a field to the paragraph.
                paragraph.AppendField("Table of Entry Field", FieldType.FieldTOCEntry);
                //Indicate whether to use table entry fields.
                tableOfContent.UseTableEntryFields = true;
                //Update the table of content.
                document.UpdateTableOfContents();
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save a Markdown file to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
    
}
