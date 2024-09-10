using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Split_by_heading
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    WordDocument newDocument = null;
                    WSection newSection = null;
                    int headingIndex = 0;
                    //Iterate each section in the Word document.
                    foreach (WSection section in document.Sections)
                    {
                        // Clone the section and add into new document.
                        if (newDocument != null)
                            newSection = AddSection(newDocument, section);
                        //Iterate each child entity in the Word document.
                        foreach (TextBodyItem item in section.Body.ChildEntities)
                        {
                            //If item is paragraph, then check for heading style and split.
                            //else, add the item into new document.
                            if (item is WParagraph)
                            {
                                WParagraph paragraph = item as WParagraph;
                                //If paragraph has Heading 1 style, then save the traversed content as separate document.
                                //And create new document for new heading content.
                                if (paragraph.StyleName == "Heading 1")
                                {
                                    if (newDocument != null)
                                    {
                                        //Saves the Word document
                                        string fileName = @"Output/Document" + (headingIndex + 1) + ".docx";
                                        SaveWordDocument(newDocument, fileName);
                                        headingIndex++;
                                    }
                                    //Create new document for new heading content.
                                    newDocument = new WordDocument();
                                    newSection = AddSection(newDocument, section);
                                    AddEntity(newSection, paragraph);
                                }
                                else if (newDocument != null)
                                    AddEntity(newSection, paragraph);
                            }
                            else
                                AddEntity(newSection, item);
                        }
                    }
                    //Save the remaining content as separate document.
                    if (newDocument != null)
                    {
                        //Saves the Word document
                        string fileName = @"Output/Document" + (headingIndex + 1) + ".docx";
                        SaveWordDocument(newDocument, fileName);
                    }
                }
            }
        }
        /// <summary>
        /// Add new section into Word document
        /// </summary>
        private static WSection AddSection(WordDocument newDocument, WSection section)
        {
            //Create new session based on original document
            WSection newSection = section.Clone();
            newSection.Body.ChildEntities.Clear();
            //Remove the first page header.
            newSection.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
            //Remove the first page footer.
            newSection.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
            //Remove the odd footer.
            newSection.HeadersFooters.OddFooter.ChildEntities.Clear();
            //Remove the odd header.
            newSection.HeadersFooters.OddHeader.ChildEntities.Clear();
            //Remove the even header.
            newSection.HeadersFooters.EvenHeader.ChildEntities.Clear();
            //Remove the even footer.
            newSection.HeadersFooters.EvenFooter.ChildEntities.Clear();
            //Add cloned section into new document
            newDocument.Sections.Add(newSection);
            return newSection;
        }
        /// <summary>
        /// Add Entity in to new section
        /// </summary>
        private static void AddEntity(WSection newSection, Entity entity)
        {
            //Add cloned item into the newly created section
            newSection.Body.ChildEntities.Add(entity.Clone());
        }
        /// <summary>
        /// Save Word document
        /// </summary>
        private static void SaveWordDocument(WordDocument newDocument, string fileName)
        {
            using (FileStream outputStream = new FileStream(Path.GetFullPath(fileName), FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                //Save file stream as Word document
                newDocument.Save(outputStream, FormatType.Docx);
                //Closes the document
                newDocument.Close();
                newDocument = null;
            }
        }
    }
}
