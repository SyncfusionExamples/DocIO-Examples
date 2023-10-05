using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.IO;

namespace Find_and_replace_text_with_TOC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    TextSelection[] selections = document.FindAll("[Insert TOC]", true, true);
                    WTextRange textrange = selections[0].GetAsOneRange();
                    WParagraph paragraph = textrange.OwnerParagraph;
                    //Remove the existing text
                    paragraph.ChildEntities.Remove(textrange);
                    //Append the TOC
                    paragraph.AppendTOC(1, 3);
                    //Update the TOC
                    document.UpdateTableOfContents();
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Data/Result_check.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
