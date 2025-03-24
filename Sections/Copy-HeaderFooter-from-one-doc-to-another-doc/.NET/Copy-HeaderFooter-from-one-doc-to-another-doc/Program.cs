using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Copy_HeaderFooter_from_one_doc_to_another_doc
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the template document stream for reading.
            using (FileStream templateStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Open the main document stream for reading.
                using (FileStream maindocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/MainDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Load the main Word document.
                    using (WordDocument mainDocument = new WordDocument(maindocumentStreamPath, FormatType.Docx))
                    {
                        // Load the template Word document.
                        using (WordDocument templateDocument = new WordDocument(templateStreamPath, FormatType.Docx))
                        {
                            // Copy header and footer from the template document to the main document.
                            MoveHeaderFooter(templateDocument, mainDocument);

                            // Create a file stream for saving the updated document.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                // Save the modified main document to the output file.
                                mainDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Move header and footer from one Word document to another Word document.
        /// </summary>
        /// <param name="templateDocument"></param>
        /// <param name="mainDocument"></param>
        public static void MoveHeaderFooter(WordDocument templateDocument, WordDocument mainDocument)
        {
            //Gets the section in the template word document
            WSection templateDocSection = templateDocument.Sections[0] as WSection;
            //Move all the items from one collection to another collection
            for (int i = 0; i < mainDocument.Sections.Count; i++)
            {
                WSection mainDocSection = mainDocument.Sections[i] as WSection;
                MoveItems(mainDocSection.HeadersFooters.EvenHeader.ChildEntities, templateDocSection.HeadersFooters.Header.ChildEntities);
                MoveItems(mainDocSection.HeadersFooters.EvenFooter.ChildEntities, templateDocSection.HeadersFooters.Footer.ChildEntities);
                MoveItems(mainDocSection.HeadersFooters.OddHeader.ChildEntities, templateDocSection.HeadersFooters.Header.ChildEntities);
                MoveItems(mainDocSection.HeadersFooters.OddFooter.ChildEntities, templateDocSection.HeadersFooters.Footer.ChildEntities);
                MoveItems(mainDocSection.HeadersFooters.FirstPageHeader.ChildEntities, templateDocSection.HeadersFooters.Header.ChildEntities);
                MoveItems(mainDocSection.HeadersFooters.FirstPageFooter.ChildEntities, templateDocSection.HeadersFooters.Footer.ChildEntities);

                //Copy the page setup from template document to main document.
                mainDocSection.PageSetup.DifferentOddAndEvenPages = templateDocSection.PageSetup.DifferentOddAndEvenPages;
                mainDocSection.PageSetup.DifferentFirstPage = templateDocSection.PageSetup.DifferentFirstPage;
                mainDocSection.PageSetup.HeaderDistance = templateDocSection.PageSetup.HeaderDistance;
                mainDocSection.PageSetup.FooterDistance = templateDocSection.PageSetup.FooterDistance;
                mainDocSection.PageSetup.Margins = templateDocSection.PageSetup.Margins;
            }
        }
        /// <summary>
        /// Move all the items from one collection to another collection.
        /// </summary>
        private static void MoveItems(EntityCollection childEntities1, EntityCollection childEntities2)
        {
            for (int i = 0; i < childEntities2.Count; i++)
                childEntities1.Add(childEntities2[i].Clone());
        }
    }
}
