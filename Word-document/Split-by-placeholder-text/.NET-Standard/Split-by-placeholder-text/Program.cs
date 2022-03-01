using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Split_a_document_by_placeholder_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document into DocIO instance.
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                String[] findPlaceHolderWord = new string[] { "[First Content Start]", "[Second Content Start]", "[Third Content Start]" };
                for (int i = 0; i < findPlaceHolderWord.Length; i++)
                {
                    //Get the start placeholder paragraph in the document.
                    WParagraph startParagraph = document.Find(findPlaceHolderWord[i], true, true).GetAsOneRange().OwnerParagraph;
                    //Get the end placeholder paragraph in the document.
                    WParagraph endParagraph = document.Find(findPlaceHolderWord[i].Replace("Start", "End"), true, true).GetAsOneRange().OwnerParagraph;
                    //Get the text body.
                    WTextBody textBody = startParagraph.OwnerTextBody;
                    //Get the start PlaceHolder index.
                    int startPlaceHolderIndex = textBody.ChildEntities.IndexOf(startParagraph);
                    //Get the end PlaceHolder index.
                    int endPlaceHolderIndex = textBody.ChildEntities.IndexOf(endParagraph);

                    //Create a new Word document.
                    WordDocument newDocument = new WordDocument();
                    newDocument.AddSection();
                    //Add the retrieved content into another new document.
                    for (int j = startPlaceHolderIndex + 1; j < endPlaceHolderIndex; j++)
                        newDocument.LastSection.Body.ChildEntities.Add(textBody.ChildEntities[j].Clone());
                    //Save the Word document to file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result" + i + ".docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        
                        newDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
