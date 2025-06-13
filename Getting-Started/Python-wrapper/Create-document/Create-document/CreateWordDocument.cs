using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.IO;
using Syncfusion.Licensing;

namespace Create_document
{
    public class CreateWordDocument
    {
        public string WordDocument(string outputPath)
        {
            try
            {
                // Create a new Word document.
                using (WordDocument document = new WordDocument())
                {
                    // Add a section and a paragraph with sample text.
                    IWSection section = document.AddSection();
                    IWParagraph paragraph = section.AddParagraph();
                    paragraph.AppendText("In 2000, AdventureWorks Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");

                    // Save the document to the specified path.
                    using (FileStream outputPathStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputPathStream, FormatType.Docx);
                    }
                }

                return "Word document created successfully at: " + outputPath;
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }

    }
}
