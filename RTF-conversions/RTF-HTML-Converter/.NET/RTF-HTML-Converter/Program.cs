using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace RTF_HTML_Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get all files from the "Data" directory.
            string[] files = Directory.GetFiles(@"Data");
            RTF_HTML_Conversion(files);
        }

        /// <summary>
        /// Converts RTF documents to HTML and vice versa.
        /// </summary>
        private static void RTF_HTML_Conversion(string[] files)
        {
            // Iterate through each input file in the given directory
            foreach (string inputFile in files)
            {
                using (FileStream inputStream = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                {
                    using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
                    {
                        // Determine the output format based on the actual format of the document
                        // If the document is RTF, convert it to HTML, otherwise convert it to RTF
                        FormatType outputFormat = (document.ActualFormatType == FormatType.Rtf) ? FormatType.Html : FormatType.Rtf;
                        // Create a destination path by changing the extension of the input file
                        string destinationPath = Path.ChangeExtension(inputFile.Replace("Data", "Output"), outputFormat.ToString());
                        using (FileStream outputStream = new FileStream(destinationPath, FileMode.Create, FileAccess.ReadWrite))
                        {
                            // Save the document in the determined output format (either HTML or RTF)
                            document.Save(outputStream, outputFormat);
                        }
                    }
                }
            }
        }
    }
}
