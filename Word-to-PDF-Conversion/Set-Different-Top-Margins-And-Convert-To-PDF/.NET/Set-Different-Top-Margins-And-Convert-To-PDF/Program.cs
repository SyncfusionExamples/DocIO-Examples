using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Set_Different_Top_Margins_And_Convert_To_PDF
{
    class Program
    {
        public static void Main(string[] args)
        {            
            //Load the Word document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx"));
            // Loop through all sections in the document
            for (int i = 0; i < document.Sections.Count; i++)
            {
                // Get the current section
                IWSection section = document.Sections[i];
                // Set the top margin based on whether it's the first or last section
                if (i == 0 || i == document.Sections.Count -1)
                    section.PageSetup.Margins.Top = 200; // Apply a top margin of 200 for the first section and last section
                else
                    section.PageSetup.Margins.Top = 90;  // Apply a top margin of 90 for all other sections
            }
            //Creates an instance of DocIORenderer.
            DocIORenderer converter = new DocIORenderer();
            // Convert the Word document to PDF
            PdfDocument pdf = converter.ConvertToPDF(document);
            //Saves the PDF file to file system.
            FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Output/Output.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
            pdf.Save(fileStream);
            // Dispose resources
            fileStream.Dispose();
            pdf.Dispose();
            converter.Dispose();
            document.Close();
        }
    }
}