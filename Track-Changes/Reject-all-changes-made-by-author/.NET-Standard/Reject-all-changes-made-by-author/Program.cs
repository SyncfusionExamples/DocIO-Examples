using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Reject_all_changes_made_by_author
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Iterates into all the revisions in Word document.
                    for (int i = document.Revisions.Count - 1; i >= 0; i--)
                    {
                        //Checks the author of current revision and rejects it.
                        if (document.Revisions[i].Author == "Nancy Davolio")
                            document.Revisions[i].Reject();
                        //Resets to last item when reject the moving related revisions.
                        if (i > document.Revisions.Count - 1)
                            i = document.Revisions.Count;
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
