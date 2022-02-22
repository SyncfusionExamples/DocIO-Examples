using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Xml_mapping
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Adds new XmlPart to the document.
                CustomXMLPart xmlPart = new CustomXMLPart(document);
                //Loads the xml code.
                xmlPart.LoadXML(@"<books><book><author>Matt Hank</author><title>New Migration Paths of the Red Breasted Robin</title><genre>New non-fiction</genre><price>29.95</price><pub_datee>12/1/2007</pub_datee> <abstract>New You see them in the spring outside your windows.</abstract></book></books>");
                //Adds text.
                paragraph.AppendText("Book author name : ");
                //Adds new content control to the paragraph.
                InlineContentControl control = paragraph.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                //Creates the XML mapping on a content control for specified XPath.
                control.ContentControlProperties.XmlMapping.SetMapping("/books/book/author", "", xmlPart);
                //Selects the single node.
                CustomXMLNode node = xmlPart.SelectSingleNode("/books/book/title");
                //Adds another paragraph.
                paragraph = section.AddParagraph();
                //Adds text.
                paragraph.AppendText("Book title: ");
                //Appends content control to second paragraph
                control = paragraph.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                //Creates the XML data mapping on a content control for specified node.
                control.ContentControlProperties.XmlMapping.SetMappingByNode(node);
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
