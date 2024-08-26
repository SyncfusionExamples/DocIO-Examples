using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace Find_and_replace_with_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Finds all the content placeholder text in the Word document.
                    TextSelection[] textSelections = document.FindAll(new Regex(@"\[(.*)\]"));
                    for (int i = 0; i < textSelections.Length; i++)
                    {
                        //Replaces the content placeholder text with desired Word document.
                        using (WordDocument subDocument = new WordDocument(new FileStream(Path.GetFullPath(@"Data/" + textSelections[i].SelectedText.TrimStart('[').TrimEnd(']') + ".docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite), FormatType.Docx))
                        {
                            document.Replace(textSelections[i].SelectedText, subDocument, true, true);
                        }
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
