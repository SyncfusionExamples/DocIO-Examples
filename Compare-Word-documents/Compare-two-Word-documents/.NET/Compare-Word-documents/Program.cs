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
                            // Create a memory stream to store the comparison result.
                            MemoryStream stream = new MemoryStream();

                            // Compare the original and revised Word documents.
                            originalWordDocument.Compare(revisedWordDocument);
							
							//Reset the stream position.
							stream.Position = 0;
							
                            //Save the stream as file.
                            using (FileStream fileStreamOutput = File.Create("Result.docx"))
                            {
                                stream.CopyTo(fileStreamOutput);
                            }
                        }                      
                    }
                }              
            }
        }       
    }
}
