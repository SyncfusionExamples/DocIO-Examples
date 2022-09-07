using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Position_for_footnote_and_endnote
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Create a section.
                IWSection section = document.AddSection();
                //Add a paragraph to a section.
                IWParagraph paragraph = section.AddParagraph();
                //Append the text to paragraph.
                paragraph.AppendText("First paragraph in First section");
                //Append the footnote.
                WFootnote footnote = paragraph.AppendFootnote(FootnoteType.Footnote) as WFootnote;
                //Set the footnote character format.
                footnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                //Set the numbering format for footnote.
                document.FootnoteNumberFormat = FootEndNoteNumberFormat.Arabic;
                //Add footnote text.
                paragraph = footnote.TextBody.AddParagraph();
                paragraph.AppendText("Footnote content");
                //Set the footnote position.
                document.FootnotePosition = FootnotePosition.PrintImmediatelyBeneathText;
                //Add the new section to the document.
                section = document.AddSection();
                //Add a paragraph to a section.
                paragraph = section.AddParagraph();
                //Append text into the paragraph.
                paragraph.AppendText("Paragraph in Second section.");
                //Append the endnote.
                WFootnote endnote = paragraph.AppendFootnote(FootnoteType.Endnote) as WFootnote;
                //Set the endnote character format.
                endnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                //Set the numbering format for endnote.
                document.EndnoteNumberFormat = FootEndNoteNumberFormat.LowerCaseRoman;
                //Add endnote text.
                paragraph = endnote.TextBody.AddParagraph();
                paragraph.AppendText("Endnote of second section");
                //Set the endnote position
                document.EndnotePosition = EndnotePosition.DisplayEndOfSection;
                //Add the new section to the document.
                section = document.AddSection();
                //Set a section break.
                section.BreakCode = SectionBreakCode.NoBreak;
                //Add a new paragraph to a section.
                paragraph = section.AddParagraph();
                //Append text into the paragraph.
                paragraph.AppendText("Paragraph in third Section.");
                //Create file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
