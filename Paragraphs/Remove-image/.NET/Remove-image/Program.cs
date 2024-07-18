using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    WTextBody textbody = document.Sections[0].Body;
                    //Iterates through the paragraphs of the textbody.
                    foreach (WParagraph paragraph in textbody.Paragraphs)
                    {
                        //Iterates through the child elements of paragraph.
                        for (int i = 0; i < paragraph.ChildEntities.Count; i++)
                        {
                            //Removes images from the paragraph.
                            if (paragraph.ChildEntities[i] is WPicture)
                            {
                                paragraph.Items.RemoveAt(i);
                                i--;
                            }
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
