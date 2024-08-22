using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Accept_or_reject_specific_type
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens the Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Iterates into all the revisions in Word document.
                    for (int i = document.Revisions.Count - 1; i >= 0; i--)
                    {
                        // Gets the type of the track changes revision.
                        RevisionType revisionType = document.Revisions[i].RevisionType;
                        //Accepts only insertion and Move from revisions changes.
                        if (revisionType == RevisionType.Insertions || revisionType == RevisionType.MoveFrom)
                            document.Revisions[i].Accept();
                        //Resets to last item when accept the moving related revisions.
                        if (i > document.Revisions.Count - 1)
                            i = document.Revisions.Count;
                    }
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
