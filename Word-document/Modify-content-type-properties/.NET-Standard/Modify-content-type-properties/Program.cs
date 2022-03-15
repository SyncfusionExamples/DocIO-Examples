using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Modify_content_type_properties
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Processes the metaproperty collection in the Word document
                    MetaProperties metaProperties = document.ContentTypeProperties;
                    //Iterates through each of the child items of metaproperties
                    for (int i = 0; i < metaProperties.Count; i++)
                    {
                        //Checks for particular display name of meta data and modifies its value
                        switch (metaProperties[i].DisplayName)
                        {
                            case "Progress Status":
                                if (metaProperties[i].Type == MetaPropertyType.Text && !metaProperties[i].IsReadOnly)
                                {
                                    metaProperties[i].Value = "Completed";
                                }
                                break;
                            case "Reviewed":
                                if (metaProperties[i].Type == MetaPropertyType.Boolean && !metaProperties[i].IsReadOnly)
                                {
                                    metaProperties[i].Value = true;
                                }
                                break;
                            case "Date":
                                if (metaProperties[i].Type == MetaPropertyType.DateTime && !metaProperties[i].IsReadOnly)
                                {
                                    metaProperties[i].Value = DateTime.UtcNow;
                                }
                                break;
                            case "Salary":
                                if ((metaProperties[i].Type == MetaPropertyType.Number ||
                                   metaProperties[i].Type == MetaPropertyType.Currency) && !metaProperties[i].IsReadOnly)
                                {
                                    metaProperties[i].Value = 12000;
                                }
                                break;
                            case "Url":
                                if (metaProperties[i].Type == MetaPropertyType.Url && !metaProperties[i].IsReadOnly)
                                {
                                    string[] value = { "https://www.syncfusion.com", "Syncfusion page" };
                                    metaProperties[i].Value = value;
                                }
                                break;
                            case "User":
                                if (metaProperties[i].Type == MetaPropertyType.User && !metaProperties[i].IsReadOnly)
                                {
                                    string[] value = { "1234", "Syncfusion" };
                                    metaProperties[i].Value = value;
                                }
                                break;
                            default:
                                break;
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
