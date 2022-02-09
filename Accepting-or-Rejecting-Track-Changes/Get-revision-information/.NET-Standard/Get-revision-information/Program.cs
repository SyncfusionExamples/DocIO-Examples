using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Get_revision_information
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
                    //Accesses the first revision in the word document.
                    Revision revision = document.Revisions[0];
                    //Gets the name of the user who made the specified tracked change.
                    string author = revision.Author;
                    //Gets the date and time that the tracked change was made.
                    DateTime dateTime = revision.Date;
                    //Gets the type of the track changes revision.
                    RevisionType revisionType = revision.RevisionType;
                    Console.WriteLine("Author : " + author);
                    Console.WriteLine("\nDate and Time : " + dateTime);
                    Console.WriteLine("\nRevision Type : " + revisionType);
                    Console.ReadKey();
                }
            }
        }
    }
}
