using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Split_by_section
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Iterate each section from Word document
                    for (int i = 0; i < document.Sections.Count; i++)
                    {
                        //Create new Word document
                        using (WordDocument newDocument = new WordDocument())
                        {
                            //Add cloned section into new Word document
                            newDocument.Sections.Add(document.Sections[i].Clone());
                            //Saves the Word document to  MemoryStream
                            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Section") + i + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                newDocument.Save(outputStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
