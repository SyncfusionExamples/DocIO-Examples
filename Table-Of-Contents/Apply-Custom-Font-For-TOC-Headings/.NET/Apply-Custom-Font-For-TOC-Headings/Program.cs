using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Apply_Custom_Font_For_TOC_Headings
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document
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
                IWTextRange textRange = paragraph.AppendText("First Chapter");
                //Apply character formatong to the text ranges.
                textRange.CharacterFormat.FontName = "Calibri";
                textRange.CharacterFormat.FontSize = 15;
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading1);
                //Gets the paragrapg style.
                IWParagraphStyle style = paragraph.GetStyle();
                style.CharacterFormat.ClearFormatting();
                //Adds the text into the paragraph.
                IWParagraph firstParagraph = section.AddParagraph();
                IWTextRange text = firstParagraph.AppendText(paraText);
                //Sets font name and size for the paragraphs.
                text.CharacterFormat.FontName = "Calibri";
                text.CharacterFormat.FontSize = 12;
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section.
                paragraph = section.AddParagraph();
                //Adds the text for the headings.
                textRange = paragraph.AppendText("Second Chapter");
                //Apply character formatong to the text ranges.
                textRange.CharacterFormat.FontName = "Verdana";
                textRange.CharacterFormat.FontSize = 15;
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading2);
                style = paragraph.GetStyle();
                style.CharacterFormat.ClearFormatting();
                //Adds the text into the paragraph.
                firstParagraph = section.AddParagraph();
                text = firstParagraph.AppendText(paraText);
                text.CharacterFormat.FontName = "Verdana";
                text.CharacterFormat.FontSize = 12;
                //Adds the section into the Word document.
                section = document.AddSection();
                //Adds the paragraph into the created section.
                paragraph = section.AddParagraph();
                //Adds the text into the headings.
                textRange = paragraph.AppendText("Third Chapter");
                //Apply character formatong to the text ranges.
                textRange.CharacterFormat.FontName = "Calibri";
                textRange.CharacterFormat.FontSize = 15;
                //Sets a built-in heading style.
                paragraph.ApplyStyle(BuiltinStyle.Heading3);
                style = paragraph.GetStyle();
                style.CharacterFormat.ClearFormatting();
                //Adds the text into the paragraph.
                firstParagraph = section.AddParagraph();
                text = firstParagraph.AppendText(paraText);
                text.CharacterFormat.FontName = "Calibri";
                text.CharacterFormat.FontSize = 12;
                //Updates the table of contents.
                document.UpdateTableOfContents();
                //Save and close the Word document.
                document.Save(Path.GetFullPath("../../../Output/Output.docx"));
            }
        }
    }
}
