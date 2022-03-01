using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Add_table_of_contents
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds the section into the Word document.
                IWSection section = document.AddSection();
                string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
                //Adds the paragraph into the created section.
                IWParagraph paragraph = section.AddParagraph();
                //Appends the TOC field with LowerHeadingLevel and UpperHeadingLevel to determines the TOC entries.
                paragraph.AppendTOC(1, 3);
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section.
                paragraph = section.AddParagraph();
                //Adds the text for the headings.
                paragraph.AppendText("First Chapter");
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading1);
                //Adds the text into the paragraph.
                section.AddParagraph().AppendText(paraText);
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section.
                paragraph = section.AddParagraph();
                //Adds the text for the headings.
                paragraph.AppendText("Second Chapter");
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading2);
                //Adds the text into the paragraph.
                section.AddParagraph().AppendText(paraText);
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section
                paragraph = section.AddParagraph();
                //Adds the text into the headings.
                paragraph.AppendText("Third Chapter");
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading3);
                //Adds the text into the paragraph.
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
