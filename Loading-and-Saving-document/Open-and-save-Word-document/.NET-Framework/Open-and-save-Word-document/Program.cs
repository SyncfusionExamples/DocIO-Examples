using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Open_and_save_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing document from file system through constructor of WordDocument class.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../HelloWorld.docx"), FormatType.Automatic))
            {
                //Appends text to the last paragraph of the document.
                document.LastParagraph.AppendText("Hello World");
                //Saves the document in file system.
                document.Save(Path.GetFullPath(@"../../Result.docx"), FormatType.Docx);
                //Closes the document.
                document.Close();
            }
        }
    }
}
