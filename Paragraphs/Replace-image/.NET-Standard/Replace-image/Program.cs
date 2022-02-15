using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    WTextBody textbody = document.Sections[0].Body;
                    //Iterates through the paragraphs of the textbody.
                    foreach (WParagraph paragraph in textbody.Paragraphs)
                    {
                        //Iterates through the child elements of paragraph.
                        foreach (ParagraphItem item in paragraph.ChildEntities)
                        {
                            if (item is WPicture)
                            {
                                WPicture picture = item as WPicture;
                                //Replaces the image.
                                if (picture.Title == "Bookmark")
                                {
                                    FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Image.png"), FileMode.Open, FileAccess.ReadWrite);
                                    picture.LoadImage(imageStream);
                                }
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
