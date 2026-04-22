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
                    using (WordDocument document = new WordDocument(sourceStreamPath, FormatType.Automatic))
                    {
                        //Opens the destination document 
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Docx))
                        {

                            WSection section = document.Sections[0] as WSection;
                            //Move all the items from one collection to another collection
                            for (int i = 0; i < destinationDocument.Sections.Count; i++)
                            {
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.EvenHeader.ChildEntities, section.HeadersFooters.Header.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.EvenFooter.ChildEntities, section.HeadersFooters.Footer.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.OddHeader.ChildEntities, section.HeadersFooters.Header.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.OddFooter.ChildEntities, section.HeadersFooters.Footer.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.FirstPageHeader.ChildEntities, section.HeadersFooters.Header.ChildEntities);
                                MoveItems(destinationDocument.Sections[i].HeadersFooters.FirstPageFooter.ChildEntities, section.HeadersFooters.Footer.ChildEntities);
                                destinationDocument.Sections[i].PageSetup.DifferentOddAndEvenPages = section.PageSetup.DifferentOddAndEvenPages;
                                destinationDocument.Sections[i].PageSetup.DifferentFirstPage = section.PageSetup.DifferentFirstPage;
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
        /// <param name="childEntities1"></param>
        /// <param name="childEntities2"></param>
        private static void MoveItems(EntityCollection childEntities1, EntityCollection childEntities2)
        {
            for (int i = 0; i < childEntities2.Count; i++)
                childEntities1.Add(childEntities2[i].Clone());
        }
    }
    }

