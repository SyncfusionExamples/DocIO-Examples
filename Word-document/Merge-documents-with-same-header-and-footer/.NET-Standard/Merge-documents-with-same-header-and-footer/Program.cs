using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Merge_documents_with_same_header_and_footer
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
                            //Get the default header and footer in the destination document.
                            EntityCollection firstDocumentHeader = destinationDocument.Sections[0].HeadersFooters.Header.ChildEntities;
                            EntityCollection firstDocumentFooter = destinationDocument.Sections[0].HeadersFooters.Footer.ChildEntities;
                            //Iterate source document.
                            foreach (WSection sourceSection in sourceDocument.Sections)
                            {
                                //Clear the header and footer of the source document.
                                sourceSection.HeadersFooters.Header.ChildEntities.Clear();
                                sourceSection.HeadersFooters.Footer.ChildEntities.Clear();
                                //Add the destination document header and footer to the source document.
                                foreach (Entity entity in firstDocumentHeader)
                                {
                                    sourceSection.HeadersFooters.Header.ChildEntities.Add(entity.Clone());
                                }
                                foreach (Entity entity in firstDocumentFooter)
                                {
                                    sourceSection.HeadersFooters.Footer.ChildEntities.Add(entity.Clone());
                                }
                                //Clone and merge the source document sections to the destination document.
                                destinationDocument.Sections.Add(sourceSection.Clone());
                            }
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
