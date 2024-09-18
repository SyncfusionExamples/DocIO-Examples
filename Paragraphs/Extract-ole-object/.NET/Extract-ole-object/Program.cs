using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Extract_ole_object
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an input Word template.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Extract the OLE object from the word document.
                    ExtractOLEObject(document);
                }
            }
        }

        /// <summary>
        /// Helper method to extract OLE object.
        /// </summary>
        private static void ExtractOLEObject(WordDocument document)
        {
            WOleObject oleObject = null;
            int oleIndex = -1;
            // Retrieving embedded object.
            foreach (WSection section in document.Sections)
            {
                foreach (WParagraph paragraph in section.Paragraphs)
                {
                    foreach (Entity entity in paragraph.ChildEntities)
                    {
                        //Checks for oleObject.
                        if (entity.EntityType == EntityType.OleObject)
                        {
                            //Gets OleObject.
                            oleObject = entity as WOleObject;
                            //Gets index of OleObject
                            oleIndex = paragraph.ChildEntities.IndexOf(oleObject);
                            //Gets ole type
                            string oleTypeStr = oleObject.ObjectType;
                            // Checks for Excel type so that file can be saved with proper extension.
                            if (oleTypeStr.Contains("Excel 2003 Worksheet") || oleTypeStr.StartsWith("Excel.Sheet.8") || (oleTypeStr.Contains("Excel Worksheet") || oleTypeStr.StartsWith("Excel.Sheet.12")))
                            {
                                if ((oleTypeStr.Contains("Excel Worksheet") || oleTypeStr.StartsWith("Excel.Sheet.12")))
                                {
                                    FileStream fstream = new FileStream(Path.GetFullPath(@"Output/Workbook" + oleObject.OleStorageName + ".xlsx"), FileMode.Create);
                                    fstream.Write(oleObject.NativeData, 0, oleObject.NativeData.Length);
                                    fstream.Flush();
                                    fstream.Close();
                                    break;
                                }
                                else
                                {
                                    FileStream fstream = new FileStream(Path.GetFullPath(@"Output/Workbook" + oleObject.OleStorageName + ".xls"), FileMode.Create);
                                    fstream.Write(oleObject.NativeData, 0, oleObject.NativeData.Length);
                                    fstream.Flush();
                                    fstream.Close();
                                    break;
                                }
                            }
                            //Checks for Word document embedded object and save them.
                            if (oleTypeStr.Contains("Word.Document"))
                            {
                                if (oleTypeStr.Contains("Word.Document.12"))
                                {
                                    FileStream fstream = new FileStream(Path.GetFullPath(@"Output/Sample" + oleObject.OleStorageName + ".docx"), FileMode.Create);
                                    fstream.Write(oleObject.NativeData, 0, oleObject.NativeData.Length);
                                    fstream.Flush();
                                    fstream.Close();
                                    break;
                                }
                                else if (oleTypeStr.Contains("Word.Document.8"))
                                {
                                    FileStream fstream = new FileStream(Path.GetFullPath(@"Output/Sample" + oleObject.OleStorageName + ".doc"), FileMode.Create);
                                    fstream.Write(oleObject.NativeData, 0, oleObject.NativeData.Length);
                                    fstream.Flush();
                                    fstream.Close();
                                    break;
                                }
                            }
                            //Checks for PDF embedded object and save them.
                            if (oleTypeStr.Contains("Acrobat Document") || oleTypeStr.StartsWith("AcroExch.Document.7") || (oleTypeStr.Contains("AcroExch.Document.11") || oleTypeStr.StartsWith("AcroExch.Document.DC")))
                            {
                                FileStream fstream = new FileStream(Path.GetFullPath(@"Output/Sample" + oleObject.OleStorageName + ".pdf"), FileMode.Create);
                                fstream.Write(oleObject.NativeData, 0, oleObject.NativeData.Length);
                                fstream.Flush();
                                fstream.Close();
                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}
