using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Set_author_and_date
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
                            originalDocument.Compare(revisedDocument, "Nancy Davolio", DateTime.Now.AddDays(-1));

                            //Save the Word document.
                            using (FileStream fileStreamOutput = File.Create("Output/Output.docx"))
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