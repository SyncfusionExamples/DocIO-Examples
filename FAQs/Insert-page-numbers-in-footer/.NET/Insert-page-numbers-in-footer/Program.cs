using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_page_numbers_in_footer
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens the Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Iterates through sections.
                    foreach (WSection sec in document.Sections)
                    {
                        IWParagraph para = sec.AddParagraph();
                        //Appends page field to the paragraph.
                        para.AppendField("footer", FieldType.FieldPage);
                        para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                        sec.PageSetup.PageNumberStyle = PageNumberStyle.Arabic;
                        //Adds paragraph to footer.
                        sec.HeadersFooters.Footer.Paragraphs.Add(para);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
