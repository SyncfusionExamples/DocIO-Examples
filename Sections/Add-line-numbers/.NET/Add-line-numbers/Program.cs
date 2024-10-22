using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_line_numbers
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
                    //Iterate each section.
                    foreach (WSection section in document.Sections)
                    {
                        //Set the line number distance from the text.
                        section.PageSetup.LineNumberingDistanceFromText = 10;
                        //Set the numbering mode.
                        section.PageSetup.LineNumberingMode = LineNumberingMode.Continuous;
                        //Set the starting line number value.
                        section.PageSetup.LineNumberingStartValue = 1;
                        //Set the increment value for line numbering.
                        section.PageSetup.LineNumberingStep = 2;
                    }
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
