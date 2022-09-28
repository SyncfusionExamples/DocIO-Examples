using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_section_break_with_page_break
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Replace the section break with page break in the Word document.
                    ReplaceSectionBreakWithPageBreak(document);
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
        /// Replace the section break with page break in the Word document.
        /// </summary>
        private static void ReplaceSectionBreakWithPageBreak(WordDocument document)
        {
            //Add page break and removes section break by moving the section items to the first section.
            while (document.Sections.Count > 1)
            {
                WSection sec = document.Sections[1];
                //Add page break in last paragraph of the section.
                (document.Sections[0].Body.AddParagraph()).AppendBreak(BreakType.PageBreak);
                //Iterate the section items in the Word document.
                foreach (Entity entity in sec.Body.ChildEntities)
                {
                    //Merge the next section to first section before removing it.
                    int lastItemIndex = document.Sections[0].Body.ChildEntities.Count;
                    document.Sections[0].Body.ChildEntities.Insert(lastItemIndex, entity.Clone());
                }
                //Remove section at the specified index from the Word document.
                document.Sections.RemoveAt(1);
            }
        }
    }
}
