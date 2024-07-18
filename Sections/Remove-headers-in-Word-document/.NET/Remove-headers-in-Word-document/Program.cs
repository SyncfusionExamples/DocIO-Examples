using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_headers_in_Word_document
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
                        HeaderFooter header;
                        //Gets even footer of current section.
                        header = section.HeadersFooters[HeaderFooterType.EvenHeader];
                        //Removes even footer.
                        header.ChildEntities.Clear();
                        //Gets odd footer of current section.
                        header = section.HeadersFooters[HeaderFooterType.OddHeader];
                        //Removes odd footer.
                        header.ChildEntities.Clear();
                        //Gets first page footer.
                        header = section.HeadersFooters[HeaderFooterType.FirstPageHeader];
                        //Removes first page footer.
                        header.ChildEntities.Clear();
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
