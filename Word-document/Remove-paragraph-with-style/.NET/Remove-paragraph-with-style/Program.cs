using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Remove_paragraph_with_style
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Processes the body contents for each section in the Word document
                    foreach (WSection section in document.Sections)
                    {
                        //Accesses the Body of section where all the contents in document are apart.
                        WTextBody sectionBody = section.Body;
                        IterateTextBody(sectionBody);
                        WHeadersFooters headersFooters = section.HeadersFooters;
                        //Consider that OddHeader and OddFooter are applied to this document.
                        //Iterates through the TextBody of OddHeader and OddFooter.
                        IterateTextBody(headersFooters.OddHeader);
                        IterateTextBody(headersFooters.OddFooter);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"OutPut/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Iterates textbody child elements.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterates through each of the child items of WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items (should be either paragraph, table or block content control) as IEntity.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                //Decides the element type by using EntityType.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Checks for particular style name and removes the paragraph from DOM.
                        if (paragraph.StyleName == "MyStyle")
                        {
                            int index = textBody.ChildEntities.IndexOf(paragraph);
                            textBody.ChildEntities.RemoveAt(index);
                        }
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells.
                        //Iterates through table's DOM.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterates to the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }

        /// <summary>
        /// Iterates table child elements.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterates the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Table cell is derived from (also a) TextBody.
                    //Reusing the code meant for iterating TextBody.
                    IterateTextBody(cell);
                }
            }
        }
    }
}
