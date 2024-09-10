using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Remove_footnotes_and_endnotes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Removes footnote and endnote from the document
                    RemoveFootNoteEndNote(document);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Remove footnote and endnote from Word document
        /// </summary>
        private static void RemoveFootNoteEndNote(WordDocument document)
        {
            foreach (WSection section in document.Sections)
                RemoveFootNoteEndNote(section.Body);
        }
        /// <summary>
        /// Remove footnote and endnote from textbody
        /// </summary>        
        private static void RemoveFootNoteEndNote(WTextBody textBody)
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
                            //Footnote and endnote are maintained in same entity type in DocIO
                            if (paragraph.ChildEntities[j] is WFootnote)
                            {
                                paragraph.ChildEntities.RemoveAt(j);
                                j--;
                            }
                        }
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells
                        //Iterates through table's DOM and and Remove footnote.
                        RemoveFootNoteEndNote(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterates to the body items of Block Content Control and Remove footnote.
                        RemoveFootNoteEndNote(blockContentControl.TextBody);
                        break;
                }
            }
        }
        /// <summary>
        /// Remove footnote and endnote from table
        /// </summary>
        private static void RemoveFootNoteEndNote(WTable table)
        {
            //Iterates the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Iterate items in cell and and Remove footnote.
                    RemoveFootNoteEndNote(cell);
                }
            }
        }
    }
}
