using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Show_Warning_for_unsupported_elements
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(@"../../../Smart Art.docx", FileMode.Open))
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
                            if (!renderer.IsCanceled)
                            {     //Saves the PDF file
                                FileStream outputFile = new FileStream("Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                                pdfDocument.Save(outputFile);
                                outputFile.Dispose();

                                PdfDocument.ClearFontCache();

                                Console.WriteLine("Success");

                                System.Diagnostics.Process process = new System.Diagnostics.Process();
                                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.pdf")
                                {
                                    UseShellExecute = true
                                };
                                process.Start();
                            }
                            else
                            {
                                Console.WriteLine("The execution stops due to the input document contains SmartArt");
                                Console.ReadKey();
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
    /// <seealso cref="IWarning" />
    public class DocumentWarning : IWarning
    {
        public bool ShowWarnings(List<WarningInfo> warningInfo)
        {
            bool isContinueConversion = true;
            foreach (WarningInfo warning in warningInfo)
            {
                //Based on WarningType enumeration, you can do your manipulation.
                //Skips the Word to PDF conversion by setting isContinueConversion value as false
                if (warning.WarningType == WarningType.SmartArt)
                    isContinueConversion = false;
            }
            return isContinueConversion;
        }
    }
}
