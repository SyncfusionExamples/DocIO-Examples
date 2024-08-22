using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_and_replace_text_with_image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Finds all the image placeholder text in the Word document.
                    TextSelection[] textSelections = document.FindAll(new Regex("^//(.*)"));
                    for (int i = 0; i < textSelections.Length; i++)
                    {
                        //Replaces the image placeholder text with desired image.
                        WParagraph paragraph = new WParagraph(document);
                        FileStream imageStream = new FileStream(Path.GetFullPath(@"Data" + textSelections[i].SelectedText + ".png"), FileMode.Open, FileAccess.ReadWrite);
                        WPicture picture = paragraph.AppendPicture(imageStream) as WPicture;
                        TextBodyPart bodyPart = new TextBodyPart(document);
                        bodyPart.BodyItems.Add(paragraph);
                        document.Replace(textSelections[i].SelectedText, bodyPart, true, true);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
