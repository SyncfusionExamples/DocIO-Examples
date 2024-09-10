using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_paragraph_formatting
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access the section in a Word document.
                    IWSection section = document.Sections[0];
                    //Access the paragraph in a Word document.
                    IWParagraph paragraph = section.Paragraphs[4];
                    //Apply paragraph formatting.
                    paragraph.ParagraphFormat.AfterSpacing = 18f;
                    paragraph.ParagraphFormat.BeforeSpacing = 18f;
                    paragraph.ParagraphFormat.BackColor = Color.LightGray;
                    paragraph.ParagraphFormat.FirstLineIndent = 10f;
                    paragraph.ParagraphFormat.LineSpacing = 10f;
                    paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                    //Access the paragraph in a Word document.
                    paragraph = section.Paragraphs[7];
                    //Apply keep lines together property to the paragraph.
                    paragraph.ParagraphFormat.Keep = true;
                    //Access the paragraph in a Word document.
                    paragraph = section.Paragraphs[6];
                    //Apply keep with next property to the paragraph.
                    paragraph.ParagraphFormat.KeepFollow = true;
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}