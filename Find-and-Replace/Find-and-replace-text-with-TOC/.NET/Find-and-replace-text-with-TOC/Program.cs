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
                    int index = paragraph.ChildEntities.IndexOf(textrange);
                    //Remove the existing text
                    paragraph.ChildEntities.Remove(textrange);
                    //Insert the TOC
                    InsertTOC(document, paragraph, index);
                    //Update the TOC
                    document.UpdateTableOfContents();
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Data/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Insert the TOC in the index of given paragraph
        /// </summary>
        /// <param name="document"></param>
        /// <param name="ownerPara"></param>
        /// <param name="index"></param>
        private static void InsertTOC(WordDocument document, WParagraph ownerPara, int index)
        {
            //Create a new paragraph
            WParagraph newPara = new WParagraph(document);
            //Append TOC to the new paragraph
            newPara.AppendTOC(1, 3);
            //Insert the child entities of new paragraph to the owner paragraph at the given index.
            for (int i = 0; i < newPara.ChildEntities.Count;)
            {
                ownerPara.ChildEntities.Insert(index, newPara.ChildEntities[i]);
                index++;
            }
        }
    }
}
