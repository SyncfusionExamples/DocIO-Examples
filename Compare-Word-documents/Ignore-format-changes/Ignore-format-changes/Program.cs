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

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the Original file as Stream.
            using (FileStream originalDocStream = new FileStream(Path.GetFullPath(@"../../../Data/OriginalDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                //Open the Revised file as Stream
                using (FileStream revisedDocStream = new FileStream(Path.GetFullPath(@"../../../Data/RevisedDocument.docx"), FileMode.Open, FileAccess.Read))
                {
                    //Loads Original file stream into Word document.
                    using (WordDocument originalWordDocument = new WordDocument(originalDocStream, FormatType.Docx))
                    {
                        //Loads Revised file stream into Word document.
                        using (WordDocument revisedWordDocument = new WordDocument(revisedDocStream, FormatType.Docx))
                        {                            
                            //Sets the Comparison option detect format changes, whether to detect format changes while comparing two Word documents.
                            ComparisonOptions compareOptions = new ComparisonOptions();
                            compareOptions.DetectFormatChanges = false;

                            //Compares the original document with revised document.
                            originalDocument.Compare(revisedDocument, "Your name", DateTime.Now, compareOptions);

                            //Creates file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Saves the Word document to file stream.
                                originalWordDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
