using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Extract_images_from_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens the Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    WTextBody textbody = document.Sections[0].Body;
                    byte[] image;
                    int i = 1;
                    //Iterates through the paragraphs.
                    foreach (WParagraph paragraph in textbody.Paragraphs)
                    {
                        //Iterates through the paragraph items. 
                        foreach (ParagraphItem item in paragraph.ChildEntities)
                        {
                            //Gets the picture and saves it into specified location.
                            switch (item.EntityType)
                            {
                                case EntityType.Picture:
                                    WPicture picture = item as WPicture;
                                    image = picture.ImageBytes;
                                    File.WriteAllBytes(@"Output/Output" + i + ".jpeg", image);
                                    i++;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }                    
                }
            }
        }
    }
}
