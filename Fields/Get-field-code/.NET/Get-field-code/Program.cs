using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Get_Field_Code
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document file stream.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Load the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    // Access the field in the first paragraph of the first section.
                    WField field = document.Sections[0].Paragraphs[0].ChildEntities[2] as WField;

                    // Get and print the field code of the merge field.
                    string fieldCode = field.FieldCode;
                    Console.WriteLine(fieldCode);
                }
            }
        }
    }
}
