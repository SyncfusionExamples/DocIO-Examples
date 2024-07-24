using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocToPDFConverter;

namespace Get_missing_fonts_for_PDF_conversion
{
    internal class Program
    {
        // List to store names of fonts that are not installed
        static List<string> fonts = new List<string>();
        static void Main(string[] args)
        {
            // Open the existing DOCX file stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the file stream into a Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    // Hook the font substitution event to detect missing fonts
                    wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;

                    // Instantiate DocToPDFConverter for Word to PDF conversion
                    using (DocToPDFConverter converter = new DocToPDFConverter())
                    {
                        // Convert Word document into PDF document
                        using (PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument))
                        {
                            // Save the PDF document to output stream
                            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../Data/Result.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                pdfDocument.Save(outputStream);
                            }
                        }
                    }
                }
            }

            // Print the fonts that are not available in machine, but used in Word document.
            if (fonts.Count > 0)
            {
                Console.WriteLine("Fonts not available in environment:");
                foreach (string font in fonts)
                    Console.WriteLine(font);
            }
            else
            {
                Console.WriteLine("Fonts used in Word document are available in environment.");
            }
            Console.ReadKey();
        }

        // Event handler for font substitution event
        static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            // Add the original font name to the list if it's not already there
            if (!fonts.Contains(args.OriginalFontName))
                fonts.Add(args.OriginalFontName);
        }
    }
}
