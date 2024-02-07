using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Find_and_replace_all
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../EnglishNumber - Copy.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    ReplaceEnglishNumberToArabic(document);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        private static void ReplaceEnglishNumberToArabic(WordDocument document)
        {
            document.Replace("0", "٠", false, false);
            document.Replace("1", "١", false, false);
            document.Replace("2", "٢", false, false);
            document.Replace("3", "٣", false, false);
            document.Replace("4", "٤", false, false);
            document.Replace("5", "٥", false, false);
            document.Replace("6", "٦", false, false);
            document.Replace("7", "٧", false, false);
            document.Replace("8", "٨", false, false);
            document.Replace("9", "٩", false, false);
        }
    }
}
