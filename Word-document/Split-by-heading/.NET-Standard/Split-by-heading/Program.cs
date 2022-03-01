using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Split_by_heading
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read))
            {
                //Load the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    inputStream.Dispose();
                    WordDocument newDocument = null;
                    WSection newSection = null;
                    int i = 0;
                    //Iterate each section from Word document
                    foreach (WSection section in document.Sections)
                    {
                        if (newDocument != null)
                            newSection = AddSection(newDocument, section);
                        foreach (TextBodyItem textbodyitem in section.Body.ChildEntities)
                        {
                            if (textbodyitem is WParagraph)
                            {
                                WParagraph para = textbodyitem as WParagraph;
                                if (para.StyleName == "Heading 1")
                                {
                                    if (newDocument != null)
                                    {
                                        //Saves the Word document
                                        using (FileStream outputStream = new FileStream(@"../../../Heading" + i + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                                        {
                                            SaveWordDocument(newDocument, outputStream);
                                        }
                                        i++;
                                    }
                                    //Create new Word document
                                    newDocument = new WordDocument();
                                    newSection = AddSection(newDocument, section);
                                    //Add cloned paragraphs into new section
                                    AddEntity(newSection, para);
                                }
                                else if (newDocument != null)
                                    //Add cloned paragraphs into new section
                                    AddEntity(newSection, para);                                
                            }
                            else                            
                                //Add cloned item into new section
                                AddEntity(newSection, textbodyitem);                                                      
                        }
                    }
                    if (newDocument != null)
                    {
                        //Saves the Word document to  MemoryStream
                        using (FileStream outputStream = new FileStream(@"../../../Heading" + i + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                        {
                            SaveWordDocument(newDocument, outputStream);
                        }
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
        private static void SaveWordDocument(WordDocument newDocument, FileStream outputStream)
        {
            //Save file stream as Word document
            newDocument.Save(outputStream, FormatType.Docx);
            //Closes the document
            newDocument.Close();
            newDocument = null;
        }
    }
}
