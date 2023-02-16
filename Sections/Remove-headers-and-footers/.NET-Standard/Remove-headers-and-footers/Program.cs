using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_headers_and_footers
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2VVhkQlFadV5JXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0djXn5ZcXVQRWVfVEA=");
            //Open the file as a stream.
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Get the headers and footers from destination document.
                    //EntityCollection firstpageheader = destinationDocument.Sections[0].HeadersFooters.FirstPageHeader.ChildEntities;
                    //EntityCollection firstpagefooter = destinationDocument.Sections[0].HeadersFooters.FirstPageFooter.ChildEntities;
                    //EntityCollection evenheader = destinationDocument.Sections[0].HeadersFooters.EvenHeader.ChildEntities;
                    //EntityCollection evenfooter = destinationDocument.Sections[0].HeadersFooters.EvenFooter.ChildEntities;
                    //EntityCollection oddheader = destinationDocument.Sections[0].HeadersFooters.OddHeader.ChildEntities;
                    //EntityCollection oddfooter = destinationDocument.Sections[0].HeadersFooters.OddFooter.ChildEntities;
                    //Open the source document as a stream.
                    using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(@"../../../SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source document.
                        using (WordDocument sourceDocument = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                        {
                            int count = 0;
                            //Iterate source document.
                            foreach (WSection sourceDocumentSection in sourceDocument.Sections)
                            {
                                
                                //Remove the first page header.
                                sourceDocumentSection.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                                //Remove the first page footer.
                                sourceDocumentSection.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                                //Remove the even header.
                                sourceDocumentSection.HeadersFooters.EvenHeader.ChildEntities.Clear();
                                //Remove the even footer.
                                sourceDocumentSection.HeadersFooters.EvenFooter.ChildEntities.Clear();
                                //Remove the odd header.
                                sourceDocumentSection.HeadersFooters.OddHeader.ChildEntities.Clear();
                                //Remove the odd footer.
                                sourceDocumentSection.HeadersFooters.OddFooter.ChildEntities.Clear();
                                //sourceDocumentSection.HeadersFooters.LinkToPrevious = true;
                                //destinationDocument.Sections.Add(sourceDocumentSection.Clone());
                                //foreach (Entity èntity in firstpageheader)
                                //    (sourceDocumentSection.HeadersFooters.FirstPageHeader as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                //foreach (Entity èntity in firstpagefooter)
                                //    (sourceDocumentSection.HeadersFooters.FirstPageFooter as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                //foreach (Entity èntity in evenheader)
                                //    (sourceDocumentSection.HeadersFooters.EvenHeader as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                //foreach (Entity èntity in evenfooter)
                                //    (sourceDocumentSection.HeadersFooters.EvenFooter as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                //foreach (Entity èntity in oddheader)
                                //    (sourceDocumentSection.HeadersFooters.OddHeader as HeaderFooter).ChildEntities.Add(èntity.Clone());
                                //foreach (Entity èntity in oddfooter)
                                //    (sourceDocumentSection.HeadersFooters.OddFooter as HeaderFooter).ChildEntities.Add(èntity.Clone());
                            }
                            destinationDocument.ImportContent(sourceDocument,ImportOptions.UseDestinationStyles);
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
