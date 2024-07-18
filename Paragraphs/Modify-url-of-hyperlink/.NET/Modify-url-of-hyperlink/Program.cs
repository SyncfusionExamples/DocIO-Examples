using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_url_of_hyperlink
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
                    WParagraph paragraph = document.LastParagraph;
                    //Iterates through the paragraph items.
                    foreach (ParagraphItem item in paragraph.ChildEntities)
                    {
                        if (item is WField)
                        {
                            if ((item as WField).FieldType == FieldType.FieldHyperlink)
                            {
                                //Gets the hyperlink field.
                                Hyperlink link = new Hyperlink(item as WField);
                                if (link.Type == HyperlinkType.WebLink)
                                {
                                    //Modifies the url of the hyperlink.
                                    link.Uri = "http://www.google.com";
                                    link.TextToDisplay = "Google";
                                    break;
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
