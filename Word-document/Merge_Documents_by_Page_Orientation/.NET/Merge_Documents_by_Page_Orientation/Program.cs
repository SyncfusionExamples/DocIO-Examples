using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Merge_Documents_by_Page_Orientation
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream sourceStream = new FileStream(Path.GetFullPath("../../../Data/MainDocument.rtf"), FileMode.Open))
            using (FileStream destinationStream = new FileStream(Path.GetFullPath("../../../Data/Template.rtf"), FileMode.Open))
            using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Automatic))
            using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Rtf))
            {
                //Iterate through each section in the source document
                foreach (WSection section in sourceDocument.Sections)
                {
                    bool skipSection = false;
                    // Check if the section has landscape orientation which contain table
                    if (section.PageSetup.Orientation == PageOrientation.Landscape)
                    {
                        for (int i = 0; i < section.Body.ChildEntities.Count; i++)
                        {
                            if (section.Body.ChildEntities[i] is WTable)
                            {
                                skipSection = true;
                                break;
                            }
                        }
                    }
                    // If the section is not marked to be skipped, clone and add it to the destination document
                    if (!skipSection)
                    {
                        destinationDocument.Sections.Add(section.Clone());
                    }
                }

                // Copy header from other document
                //If document already have an header clear the header part before copying
                using (FileStream headerStream = new FileStream(Path.GetFullPath("../../../Data/header.rtf"), FileMode.Open))
                {
                    WordDocument headerDoc = new WordDocument(headerStream, FormatType.Rtf);
                    MoveHeader(headerDoc, destinationDocument);
                    headerDoc.Close();
                }
                // Copy footer from other document
                using (FileStream footerStream = new FileStream(Path.GetFullPath("../../../Data/footer.rtf"), FileMode.Open))
                {
                    WordDocument footerDoc = new WordDocument(footerStream, FormatType.Rtf);
                    MoveFooter(footerDoc, destinationDocument);
                    footerDoc.Close();
                }

                //Saves the Word document to FileStream.
                using (FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/Result.docx"), FileMode.Create, FileAccess.Write))
                {
                    destinationDocument.Save(outputStream, FormatType.Docx);
                }
            }
        }
        // Method to move header content from the template document to the main document
        static void MoveHeader(WordDocument templateDocument, WordDocument mainDocument)
        {
            //Entity Collection 
            EntityCollection header = templateDocument.Sections[0].Body.ChildEntities;
            //Move all the items from one collection to another collection
            for (int i = 0; i < mainDocument.Sections.Count; i++)
            {
                WSection mainDocSection = mainDocument.Sections[i] as WSection;
                // Move the header items from the template to the main document's Header
                MoveItems(mainDocSection.HeadersFooters.OddHeader.ChildEntities, header);
                MoveItems(mainDocSection.HeadersFooters.EvenHeader.ChildEntities, header);
                MoveItems(mainDocSection.HeadersFooters.FirstPageHeader.ChildEntities, header);
            }
        }
        // Method to move footer content from the template document to the main document
        static void MoveFooter(WordDocument templateDocument, WordDocument mainDocument)
        {
            EntityCollection footer = templateDocument.Sections[0].Body.ChildEntities;
            //Move all the items from one collection to another collection
            for (int i = 0; i < mainDocument.Sections.Count; i++)
            {
                WSection mainDocSection = mainDocument.Sections[i] as WSection;
                // Move the footer items from the template to the main document's Footer
                MoveItems(mainDocSection.HeadersFooters.OddFooter.ChildEntities, footer);
                MoveItems(mainDocSection.HeadersFooters.EvenFooter.ChildEntities, footer);
                MoveItems(mainDocSection.HeadersFooters.FirstPageFooter.ChildEntities, footer);
            }
        }
        // Move the footer items from the template to the main document's header footer
        static void MoveItems(EntityCollection destinationDoc, EntityCollection sourceDoc)
        {
            for (int i = 0; i < sourceDoc.Count; i++)
            {
                destinationDoc.Add(sourceDoc[i].Clone());
            }
        }
    }
}
