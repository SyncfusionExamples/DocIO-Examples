using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Merge_documents_with_same_margins
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the destination document.
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                        {
                            //Access pagesetup in the destination document.
                            WPageSetup destinationPageSetup = destinationDocument.LastSection.PageSetup;
                            //Access section in the source document.
                            foreach (WSection section in sourceDocument.Sections)
                            {
                                section.PageSetup.Margins = destinationPageSetup.Margins;
                            }
                            //Import the contents of source document at the end of destination document.
                            destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);
                            //Create a file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the Word document to the file stream.
                                destinationDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
