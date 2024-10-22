using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;


namespace Save_Word_in_old_compatibility
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create an instance of WordDocument.
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                //Append paragraph.
                document.LastParagraph.AppendText("Hello World");
                //Sets the compatibility mode to Word 2007.
                document.Settings.CompatibilityMode = CompatibilityMode.Word2007;
                //Create FileStream to save the Word file.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the Word file.
                    document.Save(outputStream, FormatType.Docx);
                }
            }             
        }
    }
}
