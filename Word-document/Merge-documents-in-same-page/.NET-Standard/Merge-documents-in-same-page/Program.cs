using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Merge_documents_in_same_page
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an source document from file system through constructor of WordDocument class.
                using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Opens the destination document.
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                        {
                            //Sets the break-code of First section of source document as NoBreak to avoid imported from a new page.
                            sourceDocument.Sections[0].BreakCode = SectionBreakCode.NoBreak;
                            //Imports the contents of source document at the end of destination document.
                            destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);
                            //Creates file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
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
