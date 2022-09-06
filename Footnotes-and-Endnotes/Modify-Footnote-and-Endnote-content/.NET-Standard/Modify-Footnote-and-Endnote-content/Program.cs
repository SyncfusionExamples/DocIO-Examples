using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Modify_Footnote_and_Endnote_content
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Gets the textbody of the first section
                    WTextBody textBody = document.Sections[0].Body;
                    //Gets the paragraph at index 6
                    WParagraph paragraph = textBody.Paragraphs[6];
                    //Gets the footnote at index 0
                    WFootnote footnote = paragraph.ChildEntities[0] as WFootnote;
                    //Clear footnote content.
                    footnote.TextBody.ChildEntities.Clear();
                    //Add Paragraph to body of footnote.
                    WParagraph footnoteParagraph = footnote.TextBody.AddParagraph() as WParagraph;
                    //Sets the footnote character format.
                    footnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    //Append footnotes text.
                    footnoteParagraph.AppendText("Footnote is modified.");

                    //Gets the textbody of the third section
                    textBody = document.Sections[2].Body;
                    //Gets the paragraph at index 1
                    paragraph = textBody.Paragraphs[1];
                    //Gets the footnote at index 0
                    WFootnote endnote = paragraph.ChildEntities[0] as WFootnote;
                    //Clear footnote content.
                    endnote.TextBody.ChildEntities.Clear();
                    //Add Paragraph to body of footnote.
                    WParagraph endnoteParagraph = endnote.TextBody.AddParagraph() as WParagraph;
                    //Sets the footnote character format.
                    endnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    //Append footnotes text.
                    endnoteParagraph.AppendText("Endnote is modified.");
                    using (FileStream outputStream = new FileStream(@"../../../Sample.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            
            }
        }        
    }
}
