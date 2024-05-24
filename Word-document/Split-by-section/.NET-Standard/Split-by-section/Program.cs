using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Split_by_section
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {



                //Load the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    int fileId = 1;
                    //Iterate each section from Word document
                    foreach (WSection section in document.Sections)
                    {
                        //Create new Word document
                        using (WordDocument newDocument = new WordDocument())
                        {
                            //Add cloned section into new Word document
                            newDocument.Sections.Add(section.Clone());
                            //Saves the Word document to  MemoryStream
                            using (FileStream outputStream = new FileStream(@"../../../Section" + fileId + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                newDocument.Save(outputStream, FormatType.Docx);
                            }
                        }
                        fileId++;
                    }
                }
            }
        }
    }
}
