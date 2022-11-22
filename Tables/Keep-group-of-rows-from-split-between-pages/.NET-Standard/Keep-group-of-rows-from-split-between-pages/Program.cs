using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Keep_group_of_rows_from_split_between_pages
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access table in a Word document.
                    WTable table = document.Sections[0].Tables[0] as WTable;
                    //Iterate row collection.
                    for (int i = 6; i < table.Rows.Count - 1; i++)
                    {
                        WTableRow row = table.Rows[i];
                        Entity entity = row.Cells[0].ChildEntities[0];
                        switch (entity.EntityType)
                        {
                            case EntityType.Paragraph:
                                WParagraph paragraph = entity as WParagraph;
                                //Keep paragraph together on a page.
                                paragraph.ParagraphFormat.KeepFollow = true;
                                break;
                            case EntityType.Table:
                                // Iterate through the body item of a cell.
                                IterateTextBody((entity as WTable).Rows[0].Cells[0]);
                                break;
                            case EntityType.BlockContentControl:
                                // Iterate to the body items of Block Content Control.
                                IterateTextBody((entity as IBlockContentControl).TextBody);
                                break;
                        }
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Access the text body items.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            // Access the body items (should be either paragraph, table, or block content control).
            Entity bodyItemEntity = textBody.ChildEntities[0];
            switch (bodyItemEntity.EntityType)
            {
                case EntityType.Paragraph:
                    WParagraph paragraph = bodyItemEntity as WParagraph;
                    //Keep paragraph together on a page.
                    paragraph.ParagraphFormat.KeepFollow = true;
                    break;
                case EntityType.Table:
                    // Iterate through the body item of a cell.
                    IterateTextBody((bodyItemEntity as WTable).Rows[0].Cells[0]);
                    break;
                case EntityType.BlockContentControl:
                    // Iterate to the body items of Block Content Control.
                    IterateTextBody((bodyItemEntity as BlockContentControl).TextBody);
                    break;
            }
        }
    }
}
