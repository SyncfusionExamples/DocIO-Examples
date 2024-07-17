using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Find_and_replace_first_occurrence
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Sets to replace only the first occurrence of a particular text.
                    document.ReplaceFirst = true;
                    //Finds the first occurrence of a particular text in the document.
                    TextSelection selection = document.Find("Adventure Works", false, false);
                    //Initializes text body part.
                    TextBodyPart bodyPart = new TextBodyPart(selection);
                    //Replaces a particular text with the text body part
                    document.Replace("Adventure Works Cycles", bodyPart, false, true, true);
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
}
