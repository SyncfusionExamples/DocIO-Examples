using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

namespace Create_ink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Get the ink paragraph of the document.
                    WParagraph paragraph = document.Sections[0].Paragraphs[0];
                    //Iterates through the child elements of ink paragraph.
                    for (int i = 0; i < paragraph.ChildEntities.Count; i++)
                    {
                        //Removes the ink from the paragraph.
                        if (paragraph.ChildEntities[i] is WInk)
                        {
                            paragraph.Items.RemoveAt(i);
                            i--;
                        }
                    }
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
