using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.Collections.Generic;
using System.IO;

namespace Merge_documents_with_same_header_and_footer
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Get the headers and footers from destination document.
                    EntityCollection firstpageheader = destinationDocument.Sections[0].HeadersFooters.FirstPageHeader.ChildEntities;
                    EntityCollection firstpagefooter = destinationDocument.Sections[0].HeadersFooters.FirstPageFooter.ChildEntities;
                    EntityCollection evenheader = destinationDocument.Sections[0].HeadersFooters.EvenHeader.ChildEntities;
                    EntityCollection evenfooter = destinationDocument.Sections[0].HeadersFooters.EvenFooter.ChildEntities;
                    EntityCollection oddheader = destinationDocument.Sections[0].HeadersFooters.OddHeader.ChildEntities;
                    EntityCollection oddfooter = destinationDocument.Sections[0].HeadersFooters.OddFooter.ChildEntities;
                    //Get the Source document names from the folder.
                    string[] sourceDocumentNames = Directory.GetFiles(@"../../../Data/");
                    //Merge each source document to the destination document.
                    int count = 0;
                    foreach (string subDocumentName in sourceDocumentNames)
                    {
                        //Open the source document files as a stream.
                        using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(subDocumentName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            //Open the source documents.
                            using (WordDocument sourceDocuments = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                            {
                                //Iterate source documents.
                                foreach (WSection sourceDocumentSections in sourceDocuments.Sections)
                                {
                                    //Clear the headers and footers of the source documents.
                                    sourceDocumentSections.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.EvenHeader.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.EvenFooter.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.OddHeader.ChildEntities.Clear();
                                    sourceDocumentSections.HeadersFooters.OddFooter.ChildEntities.Clear();
                                    //Add the destination document header and footer to the source documents.
                                    foreach (Entity èntity in firstpageheader)
                                        (sourceDocumentSections.HeadersFooters.FirstPageHeader as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                    foreach (Entity èntity in firstpagefooter)
                                        (sourceDocumentSections.HeadersFooters.FirstPageFooter as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                    foreach (Entity èntity in evenheader)
                                        (sourceDocumentSections.HeadersFooters.EvenHeader as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                    foreach (Entity èntity in evenfooter)
                                        (sourceDocumentSections.HeadersFooters.EvenFooter as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                    foreach (Entity èntity in oddheader)
                                        (sourceDocumentSections.HeadersFooters.OddHeader as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                    foreach (Entity èntity in oddfooter)
                                        (sourceDocumentSections.HeadersFooters.OddFooter as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                }
                                destinationDocument.ImportContent(sourceDocuments);
                            }
                            count++;
                        }
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



