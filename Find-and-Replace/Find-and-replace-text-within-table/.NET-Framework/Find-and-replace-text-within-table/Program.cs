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
            //Loads an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Input.docx"), FormatType.Docx))
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
                //Save the document.
                document.Save(Path.GetFullPath("../../Sample.docx"), FormatType.Docx);
            }
        }
    }
}
