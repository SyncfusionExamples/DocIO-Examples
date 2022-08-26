using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Position_for_footnote_and_endnote
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
				//Creates a section.
				IWSection section = document.AddSection();
				//Adds a paragraph to a section.
				IWParagraph paragraph = section.AddParagraph();
                //Appends the text to paragraph
                paragraph.AppendText("First paragraph in First section");
                //Appends the footnotes as it sets the footnote
                WFootnote footnote = paragraph.AppendFootnote(FootnoteType.Footnote) as WFootnote;
                //Sets the footnote character format
                footnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                //Adds footnote text
                paragraph = footnote.TextBody.AddParagraph();
                paragraph.AppendText("Footnote content");
                //Sets the footnote position
                document.FootnotePosition = FootnotePosition.PrintImmediatelyBeneathText;
                //Adds the new section to the document
                section = document.AddSection();
                //Adds a paragraph to a section
                paragraph = section.AddParagraph();
                //Inserts the text into the paragraph
                paragraph.AppendText("Paragraph in Second section.");
                //Appends the endnotes. Sets the footnote or endnote
                WFootnote endnote = paragraph.AppendFootnote(FootnoteType.Endnote) as WFootnote;
                //Sets the footnote character format
                endnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                //Adds endnote text
                paragraph = endnote.TextBody.AddParagraph();
                paragraph.AppendText("Endnote of second section");
                //Sets the endnote position
                document.EndnotePosition = EndnotePosition.DisplayEndOfSection;
                //Adds the new section to the document
                section = document.AddSection();
                //Sets a section break
                section.BreakCode = SectionBreakCode.NoBreak;
                //Adds a paragraph to a section
                paragraph = section.AddParagraph();
                //Inserts the text into the paragraph
                paragraph.AppendText("Paragraph in third Section.");
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
