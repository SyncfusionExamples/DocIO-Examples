using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Compare_Word_documents
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string originalFilePath = Path.GetFullPath(@"../../../Data/OriginalDocument.docx");
            string revisedFilePath = Path.GetFullPath(@"../../../Data/RevisedDocument.docx");
            string resultFilePath = Path.GetFullPath(@"../../../Result.docx");


            using (FileStream orgDocStream = new FileStream(originalFilePath, FileMode.Open, FileAccess.Read))
            using (FileStream revisedStream = new FileStream(revisedFilePath, FileMode.Open, FileAccess.Read))
            //Open the original Word document.
            using (WordDocument originalDocument = new WordDocument(orgDocStream, FormatType.Docx))
            //Open the revised Word document.
            using (WordDocument revisedDocument = new WordDocument(revisedStream, FormatType.Docx))
            {
                //Compare original document with revised document.
                originalDocument.Compare(revisedDocument);

                //Save the output Word document.
                using (FileStream resultStream = new FileStream(resultFilePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    originalDocument.Save(resultStream, FormatType.Docx);
                }
            }
        }       
    }
}
