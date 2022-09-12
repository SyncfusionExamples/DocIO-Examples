using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_and_replace_text_within_table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Iterate through the document sections.
                    foreach (WSection section in document.Sections)
                    {
                        foreach (Entity entity in section.Body.ChildEntities)
                        {
                            if (entity.EntityType == EntityType.Table)
                            {
                                WTable table = (WTable)entity;
                                //Iterate through the rows of table.
                                foreach (WTableRow row in table.Rows)
                                {
                                    //Iterate through the cells of rows.
                                    foreach (WTableCell cell in row.Cells)
                                    {
                                        //Iterates through the paragraphs of the cell.
                                        foreach (Entity ent in cell.ChildEntities)
                                        {
                                            if (ent.EntityType == EntityType.Paragraph)
                                            {
                                                WParagraph paragraph = ent as WParagraph;
                                                //Find the selection of text inside the paragraph.
                                                TextSelection[] textSelections = document.FindAll("Suppliers", false, true);
                                                for (int i = 0; i < textSelections.Length; i++)
                                                {
                                                    //Replace the specified regular expression with a TextSelection in the paragraph.
                                                    paragraph.Replace(new Regex("^//(.*)"), textSelections[i]);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath("../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }

        }
    }
}
