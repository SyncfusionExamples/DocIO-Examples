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
                //Add a section to the Word document.
                IWSection section = document.AddSection();
                //Add a paragraph to the created section.
                IWParagraph paragraph = section.AddParagraph();
                //Append the TOC field with LowerHeadingLevel and UpperHeadingLevel to determine the TOC entries.
                TableOfContent tableOfContent = paragraph.AppendTOC(1, 3);
                //Set lower heading level for TOC.
                tableOfContent.LowerHeadingLevel = 2;
                //Set upper heading level for TOC.
                tableOfContent.UpperHeadingLevel = 5;
                //Enable a flag to use default heading styles in TOC entries.
                tableOfContent.UseHeadingStyles = true;
                //Enable a flag to show page numbers in TOC entries.
                tableOfContent.IncludePageNumbers = true;
                //Disable a flag to align page numbers after the TOC entries.
                tableOfContent.RightAlignPageNumbers = false;
                //Disable a flag to preserve TOC entries without hyperlinks.
                tableOfContent.UseHyperlinks = false;
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Append text.
                paragraph.AppendText("First ");
                //Append line break.
                paragraph.AppendBreak(BreakType.LineBreak);
                paragraph.AppendText("Chapter");
                //Enable a flag to include newline characters in TOC entries.
                tableOfContent.IncludeNewLineCharacters = true;
                //Set a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading2);
                //Add a section to the Word document.
                section = document.AddSection();
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Append text.
                paragraph.AppendText("Second ");
                paragraph.AppendText("Chapter");
                //Set a built -in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading1);
                //Add a section to the Word document.
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
                //Add the text to the new paragraph of the section.
                section.AddParagraph().AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Set outline level for paragraph.
                paragraph.ParagraphFormat.OutlineLevel = OutlineLevel.Level2;
                //Append text.
                paragraph.AppendText("Outline Level Paragraph");
                //Enable a flag to consider outline level paragraphs in TOC entries.
                tableOfContent.UseOutlineLevels = true;
                //Add a section to the Word document.
                section = document.AddSection();
                //Add a paragraph to the section.
                paragraph = section.AddParagraph();
                //Append a field to the paragraph.
                paragraph.AppendField("Table of Entry Field", FieldType.FieldTOCEntry);
                //Enable a flag to use table entry fields in TOC entries.
                tableOfContent.UseTableEntryFields = true;
                //Update the table of contents.
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
