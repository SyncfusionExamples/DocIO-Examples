using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_headers_and_footers
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an input Word template.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Iterate to each section in the Word document.
                    foreach (WSection section in document.Sections)
                    {
                        //Remove the first page header.
                        section.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                        //Remove the first page footer.
                        section.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                        //Remove the odd footer.
                        section.HeadersFooters.OddFooter.ChildEntities.Clear();
                        //Remove the odd header.
                        section.HeadersFooters.OddHeader.ChildEntities.Clear();
                        //Remove the even header.
                        section.HeadersFooters.EvenHeader.ChildEntities.Clear();
                        //Remove the even footer.
                        section.HeadersFooters.EvenFooter.ChildEntities.Clear();
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
