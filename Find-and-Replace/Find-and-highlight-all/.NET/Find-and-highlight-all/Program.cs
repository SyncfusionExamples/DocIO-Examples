using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Find_and_highlight_all
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Giant Panda.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Finds the occurrence of the Word "panda" in the document.
                    TextSelection[] textSelection = document.FindAll("panda", false, true);
                    //Iterates through each occurrence and highlights it.
                    foreach (TextSelection selection in textSelection)
                    {
                        IWTextRange textRange = selection.GetAsOneRange();
                        textRange.CharacterFormat.HighlightColor = Syncfusion.Drawing.Color.Yellow;
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
