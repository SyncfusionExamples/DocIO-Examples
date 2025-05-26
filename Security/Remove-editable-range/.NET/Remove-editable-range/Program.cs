using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.IO;

namespace Remove_editable_range
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Get the editable range by Id
                    EditableRange editableRange = document.EditableRanges.FindById("0");

                    //Remove the editable range
                    document.EditableRanges.Remove(editableRange);

                    //Creates file stream
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
