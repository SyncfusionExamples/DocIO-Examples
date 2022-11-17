using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
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
                using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Docx))
                {
                    using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the destination document.
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Docx))
                        {
                            //Access the default header and footer in a Word document.
                            EntityCollection firstDocumentHeader = sourceDocument.LastSection.HeadersFooters.Header.ChildEntities;
                            EntityCollection firstDocumentFooter = sourceDocument.LastSection.HeadersFooters.Footer.ChildEntities;
                            //Add header and footer in the destination document from source document.
                            foreach (WSection section in destinationDocument.Sections)
                            {
                                section.HeadersFooters.Header.ChildEntities.Clear();
                                section.HeadersFooters.Footer.ChildEntities.Clear();
                                foreach (Entity entity in firstDocumentHeader)
                                {
                                    section.HeadersFooters.Header.ChildEntities.Add(entity.Clone());
                                }
                                foreach (Entity entity in firstDocumentFooter)
                                {
                                    section.HeadersFooters.Footer.ChildEntities.Add(entity.Clone());
                                }
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
