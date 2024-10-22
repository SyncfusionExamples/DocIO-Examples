using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Update_fields_in_document
{
    class Program
    {
        static void Main(string[] args)
        {
              //Loads an existing Word document into DocIO instance 
              using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
              {
                  using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                  {
                      //Updates the fields present in a document.
                      document.UpdateDocumentFields(true);
                      //Creates file stream.
                      using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                      {
                          //Saves the Word document to file stream.
                          document.Save(outputFileStream, FormatType.Docx);
                      }
                  }
              }
        }
    }  
}
