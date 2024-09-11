using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Merge_documents_with_destination_margin
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the destination document.
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                        {
                            //Get the page setup of the destination document.
                            WPageSetup destinationDocumentPageSetup = destinationDocument.LastSection.PageSetup;
                            //Iterate source document.
                            foreach (WSection sourceSection in sourceDocument.Sections)
                            {
                                sourceSection.PageSetup.Margins = destinationDocumentPageSetup.Margins;
                                //Clone and merge the source document sections to the destination document.
                                destinationDocument.Sections.Add(sourceSection.Clone());
                            }
                            //Create a file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
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
