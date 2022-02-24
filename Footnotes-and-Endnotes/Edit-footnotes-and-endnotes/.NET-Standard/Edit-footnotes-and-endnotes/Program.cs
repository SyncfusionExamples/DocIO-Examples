using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Edit_footnotes_and_endnotes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Removes footnote from the document
                    RemoveFootNote(document);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(@"../../../Result.docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Remove FootNote from Word document
        /// </summary>
        /// <param name="document"></param>
        private static void RemoveFootNote(WordDocument document)
        {
            foreach (WSection section in document.Sections)
            {
                RemoveFootNote(section.Body);
            }
        }
        /// <summary>
        /// Remove FootNote from textbody
        /// </summary>
        /// <param name="textBody"></param>
        private static void RemoveFootNote(WTextBody textBody)
        {
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items as IEntity
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                //Decides the element type by using EntityType
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        for (int j = 0; j < paragraph.ChildEntities.Count; j++)
                        {
                            if (paragraph.ChildEntities[j] is WFootnote)
                            {
                                paragraph.ChildEntities.RemoveAt(j);
                            }
                        }
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells
                        //Iterates through table's DOM and and Remove footnote.
                        RemoveFootNote(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterates to the body items of Block Content Control and Remove footnote.
                        RemoveFootNote(blockContentControl.TextBody);
                        break;
                }
            }
        }
        /// <summary>
        /// Remove FootNote table
        /// </summary>
        /// <param name="table"></param>
        private static void RemoveFootNote(WTable table)
        {
            //Iterates the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Iterate items in cell and and Remove footnote.
                    RemoveFootNote(cell);
                }
            }
        }
    }
}
