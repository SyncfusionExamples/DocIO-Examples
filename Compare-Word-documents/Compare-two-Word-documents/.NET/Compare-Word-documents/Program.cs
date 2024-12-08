using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Compare_Word_documents
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Loads the original document.
            using (FileStream originalDocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/OriginalDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument originalDocument = new WordDocument(originalDocumentStreamPath, FormatType.Docx))
                {
                    //Loads the revised document
                    using (FileStream revisedDocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/RevisedDocument.docx"), FileMode.Open, FileAccess.Read))
                    {
                        using (WordDocument revisedDocument = new WordDocument(revisedDocumentStreamPath, FormatType.Docx))
                        {
                            // Compare the original and revised Word documents.
                            originalDocument.Compare(revisedDocument);

                            //Save the Word document.
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath("Output/Result.docx")))
                            {
                                originalDocument.Save(fileStreamOutput, FormatType.Docx);
                            }
                        }
                    }                 
                }                           
            }
        }       
    }
}
