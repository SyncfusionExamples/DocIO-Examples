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
            using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                {
                    //Get the Source document names from the folder.
                    string[] sourceDocumentNames = Directory.GetFiles(@"../../../Data/");
                    foreach (string subDocumentName in sourceDocumentNames)
                    {
                        //Open the source document files as a stream.
                        using (FileStream sourceDocumentPathStream = new FileStream(Path.GetFullPath(subDocumentName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            //Open the source documents.
                            using (WordDocument sourceDocuments = new WordDocument(sourceDocumentPathStream, FormatType.Docx))
                            {
                                //Iterate sub document sections.
                                foreach (WSection section in sourceDocuments.Sections)
                                {
                                    //Check whether the source document having empty header footer
                                    //If empty header footer, add empty paragrah into it.
                                    if (IsEmptyHeaderFooter(section))
                                        RemoveEmptyHeaderFooter(section);
                                    //Iterate and check if header footer contains PAGE field and NUMPAGE field.
                                    //If PAGE field, set the RestartPageNumbering and PageStartingNumber API’s.
                                    //If NUMPAGE field, remove the numpage field and add text based on total number of pages.
                                    else
                                        IterateHeaderFooter(section);
                                }
                                //Import the contents of sub document at the end of main document.
                                destinationDocument.ImportContent(sourceDocuments, ImportOptions.KeepSourceFormatting);
                            }                           
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
        /// <summary>
        /// Iterate through source document section.
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
        /// Remove headers and footers.
        /// </summary>
        private static void RemoveEmptyHeaderFooter(WSection section)
        {
            //Remove existing header and footer.
            //Set empty paragraph as header and footer content.
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
        /// Check if header and footer is empty.
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
        /// Iterate through the textBody items.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterate through each of the child items of WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //Accesses the body items (should be either paragraph, table or block content control) as IEntity.
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
                        //Iterate through table's DOM.
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
        /// Iterate through the table.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterate the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Reuse the code meant for iterating TextBody.
                    IterateTextBody(cell);
                }
            }
        }
        /// <summary>
        /// Iterate through the paragraph.
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
                        //Check PAGE field.
                        if (field.FieldType == FieldType.FieldPage)
                        {
                            //Get the owner section of PAGE field.
                            WSection section = GetOwnerEntity(field) as WSection;
                            //Set the restart page numbering.
                            if (section != null && !section.PageSetup.RestartPageNumbering)
                            {
                                section.PageSetup.RestartPageNumbering = true;
                                section.PageSetup.PageStartingNumber = 1;
                            }
                        }
                        //Check NUMPAGE field.
                        else if (field.FieldType == FieldType.FieldNumPages)
                        {
                            RemoveNumPageField(field);
                        }
                        break;
                    case EntityType.InlineContentControl:
                        //Iterates to the paragraph items of inline content control.
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        IterateParagraph(inlineContentControl.ParagraphItems);
                        break;
                }
            }
        }
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
        /// Remove the NUMPage field.
        /// </summary>
        private static void RemoveNumPageField(WField field)
        {
            WParagraph paragraph = field.OwnerParagraph;
            int itemIndex = paragraph.ChildEntities.IndexOf(field);
            WTextRange textRange = new WTextRange(paragraph.Document);
            field.Document.UpdateWordCount();
            //Get the text from hyperlink field.
            textRange.Text = field.Document.BuiltinDocumentProperties.PageCount.ToString();
            textRange.ApplyCharacterFormat(field.CharacterFormat);
            //Remove the hyperlink field.
            paragraph.ChildEntities.RemoveAt(itemIndex);
            //Insert the hyperlink text.
            paragraph.ChildEntities.Insert(itemIndex, textRange);
        }
    }   
}
