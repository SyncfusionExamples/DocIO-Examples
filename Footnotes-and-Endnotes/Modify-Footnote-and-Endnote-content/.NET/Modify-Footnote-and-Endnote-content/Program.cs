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
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access paragraph in a Word document.
                    WParagraph paragraph = document.Sections[0].Paragraphs[6] as WParagraph;
                    //Access the footnote in the paragraph.
                    WFootnote footnote = paragraph.ChildEntities[0] as WFootnote;
                    //Clear the footnote content.
                    footnote.TextBody.ChildEntities.Clear();
                    //Add a new paragraph to the body of the footnote.
                    WParagraph footnoteParagraph = footnote.TextBody.AddParagraph() as WParagraph;
                    //Set the footnote character format.
                    footnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    //Append the footnote text.
                    footnoteParagraph.AppendText(" Footnote is modified.");
                    //Access paragraph in a Word document.
                    paragraph = document.Sections[2].Paragraphs[1] as WParagraph;
                    //Access the endnote in the paragraph.
                    WFootnote endnote = paragraph.ChildEntities[0] as WFootnote;
                    //Clear the endnote content.
                    endnote.TextBody.ChildEntities.Clear();
                    //Add a new paragraph to the body of the endnote.
                    WParagraph endnoteParagraph = endnote.TextBody.AddParagraph() as WParagraph;
                    //Set the endnote character format.
                    endnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    //Append the endnote text.
                    endnoteParagraph.AppendText(" Endnote is modified.");
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        //Save a Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }            
            }
        }        
    }
}
