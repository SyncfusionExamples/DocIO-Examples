using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_line_break_with_paragraph_mark
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input2.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Replace line break with paragraph mark in the Word document.
                    ReplaceLineBreakWithPara(document);
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Replace the line break with paragraph mark in the Word document.
        /// </summary>
        private static void ReplaceLineBreakWithPara(WordDocument document)
        {
            foreach (WSection section in document.Sections)
            {
                //Access the body section of Word document.
                WTextBody sectionBody = section.Body;
                IterateTextBody(sectionBody);
                WHeadersFooters headersFooters = section.HeadersFooters;
                //Iterate through the TextBody of OddHeader and OddFooter.
                IterateTextBody(headersFooters.OddHeader);
                IterateTextBody(headersFooters.OddFooter);
            }
        }

        /// <summary>
        /// Iterate textbody child elements.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterate the child items of text body.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            { 
                //Check the text body items.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //Check the entity type.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Iterate through the paragraph.
                        IterateParagraph(paragraph.Items);
                        break;
                    case EntityType.Table:
                        //Iterate through table.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterate to the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }
        /// <summary>
        /// Iterate table child elements.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterate the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    IterateTextBody(cell);
                }
            }
        }

        /// <summary>
        /// Iterate paragraph child elements.
        /// </summary>
        private static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //Check the element type by using EntityType.
                switch (entity.EntityType)
                {
                    case EntityType.Break:
                        Break breakItem = entity as Break;
                        //Replace line break with paragraph mark.
                        if (breakItem.BreakType == BreakType.LineBreak)
                        {
                            WParagraph ownerPara = breakItem.OwnerParagraph;
                            int breakIndex = ownerPara.ChildEntities.IndexOf(breakItem);
                            int paraIndex = ownerPara.OwnerTextBody.ChildEntities.IndexOf(ownerPara);

                            //Create new paragraph by cloning the existing paragraph.
                            WParagraph newPara = ownerPara.Clone() as WParagraph;
                            //Remove the child items after the line break from the old paragraph including line break.
                            for (int j = breakIndex; j < ownerPara.ChildEntities.Count;)
                            {
                                ownerPara.ChildEntities.RemoveAt(j);
                            }
                            int newParaItemsCount = ownerPara.ChildEntities.Count;
                            //Remove the child items before the line break from the new paragraph including line break.
                            while (newParaItemsCount + 1 != 0)
                            {
                                newPara.ChildEntities.RemoveAt(0);
                                newParaItemsCount--;
                            }
                            //Insert the new paragraph next to the line break paragraph.
                            ownerPara.OwnerTextBody.ChildEntities.Insert(paraIndex + 1, newPara);
                        }
                        break;
                    case EntityType.TextBox:
                        //Iterate the body items of textbox.
                        WTextBox textBox = entity as WTextBox;
                        IterateTextBody(textBox.TextBoxBody);
                        break;
                    case EntityType.Shape:
                        //Iterate the body items of shape.
                        Shape shape = entity as Shape;
                        IterateTextBody(shape.TextBody);
                        break;
                    case EntityType.InlineContentControl:
                        //Iterate the paragraph items of inline content control.
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        IterateParagraph(inlineContentControl.ParagraphItems);
                        break;
                }
            }
        }
    }
}
