using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Merge_without_changing_page_numbers
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Get the Source document names from the folder.
                    string[] sourceDocumentNames = Directory.GetFiles(Path.GetFullPath(@"Data/SourceDocuments/"));
                    foreach (string subDocumentName in sourceDocumentNames)
                    {
                        //Open the source document files as a stream.
                        using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(subDocumentName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            //Open the source documents.
                            using (WordDocument sourceDocument = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                            {
                                //Iterate sub-document sections.
                                foreach (WSection section in sourceDocument.Sections)
                                {
                                    //Check whether the source document has an empty header and footer.
                                    //If it has an empty header and footer, add an empty paragraph.
                                    if (IsEmptyHeaderFooter(section))
                                        RemoveEmptyHeaderFooter(section);
                                    //Iterate and check if the header and footer contain the PAGE and NUMPAGES fields.
                                    //If PAGE field, set the RestartPageNumbering and PageStartingNumber API’s.
                                    //If NUMPAGES field, remove the NUMPAGES field and add text based on the total number of pages.
                                    else
                                        IterateHeaderFooter(section);
                                }
                                //Import the contents of the sub-document at the end of the main document.
                                destinationDocument.ImportContent(sourceDocument, ImportOptions.KeepSourceFormatting);
                            }
                        }
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save a Word document to the file stream.
                        destinationDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Iterate through the source document section.
        /// </summary>
        private static void IterateHeaderFooter(WSection section)
        {
            WHeadersFooters headersFooters = section.HeadersFooters;
            //Iterate through the TextBody of all Headers and Footers.
            IterateTextBody(headersFooters.OddHeader);
            IterateTextBody(headersFooters.OddFooter);
            IterateTextBody(headersFooters.EvenHeader);
            IterateTextBody(headersFooters.EvenFooter);
            IterateTextBody(headersFooters.FirstPageHeader);
            IterateTextBody(headersFooters.FirstPageFooter);
        }
        /// <summary>
        /// Check if the header and footer are empty.
        /// </summary>
        private static bool IsEmptyHeaderFooter(WSection section)
        {
            if (section.HeadersFooters.OddHeader.ChildEntities.Count > 0
                || section.HeadersFooters.EvenHeader.ChildEntities.Count > 0
                || section.HeadersFooters.FirstPageHeader.ChildEntities.Count > 0
                || section.HeadersFooters.OddFooter.ChildEntities.Count > 0
                || section.HeadersFooters.EvenFooter.ChildEntities.Count > 0
                || section.HeadersFooters.FirstPageFooter.Count > 0)
                return false;
            return true;
        }
        /// <summary>
        /// Remove headers and footers and add empty paragraphs.
        /// </summary>
        private static void RemoveEmptyHeaderFooter(WSection section)
        {
            //Remove existing header and footer.
            //Set an empty paragraph as a header and footer content.
            section.HeadersFooters.OddHeader.ChildEntities.Clear();
            section.HeadersFooters.OddHeader.AddParagraph();
            section.HeadersFooters.EvenHeader.ChildEntities.Clear();
            section.HeadersFooters.EvenHeader.AddParagraph();
            section.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
            section.HeadersFooters.FirstPageHeader.AddParagraph();
            section.HeadersFooters.OddFooter.ChildEntities.Clear();
            section.HeadersFooters.OddFooter.AddParagraph();
            section.HeadersFooters.EvenFooter.ChildEntities.Clear();
            section.HeadersFooters.EvenFooter.AddParagraph();
            section.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
            section.HeadersFooters.FirstPageFooter.AddParagraph();
        }
        /// <summary>
        /// Iterate through the textBody items of a Word document.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterate through each of the child items of the WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //Accesses the body items (should be either paragraph, table, or block content control) as IEntity.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //Get the element type by using EntityType.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Iterate through the paragraph's DOM.
                        IterateParagraph(paragraph.Items);
                        break;
                    case EntityType.Table:
                        //Iterate through the table's DOM.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterate through the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }
        /// <summary>
        /// Iterate through the table in a Word document.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterate the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Reuse the code meant for iterating the TextBody.
                    IterateTextBody(cell);
                }
            }
        }
        /// <summary>
        /// Iterate through the paragraph in a Word document.
        /// </summary>
        private static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //Get the element type by using EntityType.
                switch (entity.EntityType)
                {
                    case EntityType.Field:
                        WField field = entity as WField;
                        //Check the PAGE field.
                        if (field.FieldType == FieldType.FieldPage)
                        {
                            //Get the owner section of the PAGE field.
                            WSection section = GetOwnerEntity(field) as WSection;
                            //Set the restart page numbering.
                            if (section != null && !section.PageSetup.RestartPageNumbering)
                            {
                                section.PageSetup.RestartPageNumbering = true;
                                section.PageSetup.PageStartingNumber = 1;
                            }
                        }
                        //Check the NUMPAGES field.
                        else if (field.FieldType == FieldType.FieldNumPages)
                        {
                            RemoveNumPageField(field);
                        }
                        break;
                }
            }
        }
        /// <summary>
        /// Get the entity owner.
        /// </summary>
        private static Entity GetOwnerEntity(WField field)
        {
            Entity baseEntity = field.Owner;

            while (!(baseEntity is WSection))
            {
                if (baseEntity is null)
                    return baseEntity;
                baseEntity = baseEntity.Owner;
            }
            return baseEntity;
        }
        /// <summary>
        /// Remove the NUMPAGES field.
        /// </summary>
        private static void RemoveNumPageField(WField field)
        {
            WParagraph paragraph = field.OwnerParagraph;
            int itemIndex = paragraph.ChildEntities.IndexOf(field);
            WTextRange textRange = new WTextRange(paragraph.Document);
            field.Document.UpdateWordCount();
            //Get the text from the hyperlink field.
            textRange.Text = field.Document.BuiltinDocumentProperties.PageCount.ToString();
            textRange.ApplyCharacterFormat(field.CharacterFormat);
            //Remove the hyperlink field.
            paragraph.ChildEntities.RemoveAt(itemIndex);
            //Insert the hyperlink text.
            paragraph.ChildEntities.Insert(itemIndex, textRange);
        }
    }
}
