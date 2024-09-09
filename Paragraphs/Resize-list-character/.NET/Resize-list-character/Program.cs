using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Resize_list_character
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Iterate through each section of the Word document.
                    foreach (WSection section in document.Sections)
                    {
                        //Access the Body of section where all the contents in document are apart.
                        WTextBody sectionBody = section.Body;
                        IterateTextBody(sectionBody);
                        //Iterate through the headers and footers.
                        IterateHeaderFooter(section);

                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Iterate through the headers and footers.
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
        /// Iterate through document textbody.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterate through each of the child items of WTextBody
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //Access the body items (should be either paragraph, table or block content control) as IEntity
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //Decide the element type by using EntityType
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Change the list character size.
                        if (paragraph.ListFormat != null && paragraph.ListFormat.CurrentListLevel != null)
                            paragraph.ListFormat.CurrentListLevel.CharacterFormat.FontSize = 25;
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
        /// Iterate through document table.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterate the row collection in a table
            foreach (WTableRow row in table.Rows)
            {
                //Iterate the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    IterateTextBody(cell);
                }
            }
        }
    }
}
