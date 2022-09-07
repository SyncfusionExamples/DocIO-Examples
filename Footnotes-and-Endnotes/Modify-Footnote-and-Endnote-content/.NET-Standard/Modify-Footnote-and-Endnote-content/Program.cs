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
            using (FileStream inputStream = new FileStream(@"../../../Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load file stream into Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access paragraph in Word document.
                    WParagraph paragraph = document.Sections[0].Paragraphs[6] as WParagraph;
                    //Get the footnote at index 0.
                    WFootnote footnote = paragraph.ChildEntities[0] as WFootnote;
                    //Clear footnote content.
                    footnote.TextBody.ChildEntities.Clear();
                    //Add new paragraph to body of footnote.
                    WParagraph footnoteParagraph = footnote.TextBody.AddParagraph() as WParagraph;
                    //Set the footnote character format.
                    footnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    //Append footnote text.
                    footnoteParagraph.AppendText(" Footnote is modified.");
                    //Access paragraph in Word document.
                    paragraph = document.Sections[2].Paragraphs[1] as WParagraph;
                    //Get the endnote at index 0.
                    WFootnote endnote = paragraph.ChildEntities[0] as WFootnote;
                    //Clear endnote content.
                    endnote.TextBody.ChildEntities.Clear();
                    //Add new paragraph to body of endnote.
                    WParagraph endnoteParagraph = endnote.TextBody.AddParagraph() as WParagraph;
                    //Set the endnote character format.
                    endnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                    //Append endnote text.
                    endnoteParagraph.AppendText(" Endnote is modified.");
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
