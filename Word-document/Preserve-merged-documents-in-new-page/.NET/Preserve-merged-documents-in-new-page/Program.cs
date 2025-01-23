using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Reflection.Metadata;

namespace Preserve_merged_documents_in_new_page
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an source document from file system through constructor of WordDocument class.
                using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Opens the destination document.
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                        {
                            //Sets the break-code of first section of source document 
                            sourceDocument.Sections[0].BreakCode = SectionBreakCode.NewPage;
                            //Imports the contents of the source document to the destination document, and
                            //applies the formatting of surrounding content to the destination document.
                            destinationDocument.ImportContent(sourceDocument, ImportOptions.MergeFormatting);
                            //Creates file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Saves the Word document to file stream.
                                destinationDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
