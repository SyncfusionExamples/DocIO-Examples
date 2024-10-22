using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_document_before_the_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
                {
                    using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the destination document.
                        using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
                        {
                            //Find the placeholder text in the destination document.
                            TextSelection textSelections = destinationDocument.Find("Product Overview", false, true);
                            WTextRange textRange = textSelections.GetAsOneRange();
                            //Get the paragraph index of the placeholder text. 
                            WParagraph paragraph = textRange.OwnerParagraph;
                            int index = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);
                            //Iterate source document.
                            foreach (WSection sourceSection in sourceDocument.Sections)
                            {
                                foreach (Entity entity in sourceSection.Body.ChildEntities)
                                {
                                    //Insert source document before the placeholder text.
                                    textRange.OwnerParagraph.OwnerTextBody.ChildEntities.Insert(index, entity.Clone());
                                    index++;
                                }

                            }
                            //Create a file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the Word document to the file stream.
                                destinationDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
