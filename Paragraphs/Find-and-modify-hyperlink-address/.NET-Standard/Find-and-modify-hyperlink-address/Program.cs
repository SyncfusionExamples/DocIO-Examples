using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Find_and_modify_hyperlink_address
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    for (int i = 0; i < document.Sections[0].Paragraphs.Count; i++)
                    {
                        WParagraph paragraph = document.Sections[0].Paragraphs[i];
                        //Iterate through the paragraph items.
                        foreach (ParagraphItem item in paragraph.ChildEntities)
                        {
                            if (item is WField)
                            {
                                if ((item as WField).FieldType == FieldType.FieldHyperlink)
                                {
                                    //Get the hyperlink field.
                                    Hyperlink link = new Hyperlink(item as WField);
                                    if (link.Type == HyperlinkType.WebLink && link.TextToDisplay == "support")
                                    {
                                        //Modify the url of the hyperlink.
                                        link.Uri = "http://support.syncfusion.com/";
                                    }
                                    else if (link.Type == HyperlinkType.EMailLink && link.TextToDisplay == "Email")
                                    {
                                        //Modify the url of the hyperlink.
                                        link.Uri = "sales@syncfusion.com";
                                    }
                                }
                            }
                        }
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
