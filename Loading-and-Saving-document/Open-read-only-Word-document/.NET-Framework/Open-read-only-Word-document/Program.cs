using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Open_read_only_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an empty WordDocument instance.
            using (WordDocument document = new WordDocument())
            {
                //Loads or opens an existing word document using read only stream.
                document.OpenReadOnly(Path.GetFullPath(@"../../HelloWorld.docx"), FormatType.Docx);
                //Appends text to the last paragraph of the document.
                document.LastParagraph.AppendText("Hello World");
                //Saves the document in file system.
                document.Save(Path.GetFullPath(@"../../Result.docx"), FormatType.Docx);
            }
        }
    }
}
