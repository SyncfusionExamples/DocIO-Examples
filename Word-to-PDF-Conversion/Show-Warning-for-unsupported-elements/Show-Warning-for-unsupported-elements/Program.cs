using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Show_Warning_for_unsupported_elements
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(@"../../../Input.docx", FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        renderer.Settings.Warning = new DocumentWarning();
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            if (renderer.IsCanceled)
                            {
                                Console.WriteLine("The execution stops due to the input document contains unsupported element");
                                Console.ReadKey();
                            }
                            else
                            {
                                //Saves the PDF file
                                FileStream outputFile = new FileStream("Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                                pdfDocument.Save(outputFile);
                                outputFile.Dispose();

                                Console.WriteLine("Success");
                            }
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// DocumentWarning class implements the IWarning interface
    /// </summary>
    public class DocumentWarning : IWarning
    {
        public bool ShowWarnings(List<WarningInfo> warningInfo)
        {
            bool isContinueConversion = true;
            foreach (WarningInfo warning in warningInfo)
            {
                //Based on the WarningType enumeration, you can do your manipulation.
                //Skip the Word to PDF conversion by setting the isContinueConversion value to false.
                //To stop execution if the input document has a SmartArt.
                if (warning.WarningType == WarningType.SmartArt)
                    isContinueConversion = false;

                //Warning messages for unsupported elements in the input document.
                Console.WriteLine("The input document contains " + warning.WarningType + " unsupported element.");
            }
            return isContinueConversion;
        }
    }

}
