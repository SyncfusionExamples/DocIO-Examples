using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Import_Headers_and_Footers_from_Another_WordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath("Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath("Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //Opens an source document from file system through constructor of WordDocument class
                    using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
                    {
                        //Opens the destination document 
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Docx))
                        {

                            WSection sourceDocumentSection = sourceDocument.Sections[0] as WSection;
                            //Move all the items from one collection to another collection
                            for (int i = 0; i < destinationDocument.Sections.Count; i++)
                            {
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.EvenHeader.ChildEntities, sourceDocumentSection.HeadersFooters.Header.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.EvenFooter.ChildEntities, sourceDocumentSection.HeadersFooters.Footer.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.OddHeader.ChildEntities, sourceDocumentSection.HeadersFooters.Header.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.OddFooter.ChildEntities, sourceDocumentSection.HeadersFooters.Footer.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.FirstPageHeader.ChildEntities, sourceDocumentSection.HeadersFooters.Header.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.FirstPageFooter.ChildEntities, sourceDocumentSection.HeadersFooters.Footer.ChildEntities);
                                destinationDocument.Sections[i].PageSetup.DifferentOddAndEvenPages = sourceDocumentSection.PageSetup.DifferentOddAndEvenPages;
                                destinationDocument.Sections[i].PageSetup.DifferentFirstPage = sourceDocumentSection.PageSetup.DifferentFirstPage;
                            }
                            using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create))
                            {
                                destinationDocument.Save(outputStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Move all the items from one collection to another collection.
        /// </summary>
        /// <param name="detinationDocumentEntityCollection">Destination collection</param>
        /// <param name="sourceDocumentEntityCollection">Source collection</param>
        private static void MoveItems(EntityCollection detinationDocumentEntityCollection, EntityCollection sourceDocumentEntityCollection)
        {
            for (int i = 0; i < sourceDocumentEntityCollection.Count; i++)
                detinationDocumentEntityCollection.Add(sourceDocumentEntityCollection[i].Clone());
        }
    }
}