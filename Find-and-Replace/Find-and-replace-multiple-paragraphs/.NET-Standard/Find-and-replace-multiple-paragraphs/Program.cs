using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace Find_and_replace_multiple_paragraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    using (WordDocument subDocument = new WordDocument(new FileStream(Path.GetFullPath(@"../../../Data/Source.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite), FormatType.Docx))
                    {
                        //Gets the content from another Word document.
                        TextBodyPart replacePart = new TextBodyPart(subDocument);
                        foreach (TextBodyItem bodyItem in subDocument.LastSection.Body.ChildEntities)
                        {
                            replacePart.BodyItems.Add(bodyItem.Clone());
                        }
                        string placeholderText = "PlaceHolderStart:" + "Suppliers/Vendors of Northwind" + "Customers of Northwind" + "Employee details of Northwind traders" + "The product information" + "The inventory details" + "The shippers" + "Purchase Order transactions" + "Sales Order transaction" + "Inventory transactions" + "Invoices" + "PlaceHolderEnd";
                        //Finds the text that extends to several paragraphs and replaces it with desired content.
                        document.ReplaceSingleLine(placeholderText, replacePart, false, false);
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
}
