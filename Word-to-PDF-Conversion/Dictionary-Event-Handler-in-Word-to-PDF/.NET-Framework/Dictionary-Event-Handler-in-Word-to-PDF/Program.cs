using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;

namespace Dictionary_Event_Handler_in_Word_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Data/Template.docx"), FormatType.Docx))
            {
                //Initializes the dictionary event to perform adding of dictionary during Word to PDF conversion.
                wordDocument.Hyphenator.AddDictionary += new AddDictionaryEventHandler(AddDictionary);
                //Initialize the ChartToImageConverter for converting charts during Word to pdf conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                FileStream dictionaryStream = new FileStream(Path.GetFullPath(@"../../Data/hyph_en_US.dic"), FileMode.Open);
                {
                    //Adds the hyphenation dictionary of the specified language.
                    Hyphenator.Dictionaries.Add("en-US", dictionaryStream);
                }
                //Create an instance of DocToPDFConverter.
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    //Convert Word document into PDF document.
                    using (PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument))
                    {
                        //Save the PDF file.
                        pdfDocument.Save(Path.GetFullPath(@"../../WordtoPDF.pdf"));
                    }
                }
            }
        }

        private static void AddDictionary(object sender, AddDictionaryEventArgs args)
        {
            //Sets the alternate language code when a specified language dictionary doesn’t exist in the collection.
            if (args.LanguageCode == "de-CH")
            {
                FileStream LanguagedictionaryStream = new FileStream(Path.GetFullPath(@"../../Data/hyph_de_CH.dic"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                args.DictionaryStream = LanguagedictionaryStream;
            }
        }
    }
}
