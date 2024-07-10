using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Extract_comments
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open an existing document
            using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Iterate through the comments in the Word document
                    foreach (WComment comment in document.Comments)
                    {
                        //Traverse each paragraph from the comment body items
                        foreach (WParagraph paragraph in comment.TextBody.Paragraphs)
                            Console.WriteLine(paragraph.Text);
                    }
                }
            }
        }
    }
}
