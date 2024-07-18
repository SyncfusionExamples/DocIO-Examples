using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_footers_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens the Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Iterates through the sections.
                    foreach (WSection section in document.Sections)
                    {
                        HeaderFooter footer;
                        //Gets even footer of current section.
                        footer = section.HeadersFooters[HeaderFooterType.EvenFooter];
                        //Removes even footer.
                        footer.ChildEntities.Clear();
                        //Gets odd footer of current section.
                        footer = section.HeadersFooters[HeaderFooterType.OddFooter];
                        //Removes odd footer.
                        footer.ChildEntities.Clear();
                        //Gets first page footer.
                        footer = section.HeadersFooters[HeaderFooterType.FirstPageFooter];
                        //Removes first page footer.
                        footer.ChildEntities.Clear();
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
