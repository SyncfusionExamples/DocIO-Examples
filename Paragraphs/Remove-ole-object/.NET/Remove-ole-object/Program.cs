using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_ole_object
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an input Word template.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Remove OLE object from the document.
                    RemoveOLEObject(document);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Helper method to remove OLE object.
        /// </summary>
        private static void RemoveOLEObject(WordDocument document)
        {
            bool isFieldStart = false;
            //Retrieving embedded object.
            foreach (WSection section in document.Sections)
            {
                foreach (WParagraph paragraph in section.Paragraphs)
                {
                    for (int i = 0; i < paragraph.ChildEntities.Count; i++)
                    {
                        Entity entity = paragraph.ChildEntities[i];
                        //Checks for oleObject.
                        if (entity.EntityType == EntityType.OleObject)
                        {
                            paragraph.ChildEntities.Remove(entity);
                            isFieldStart = true;
                            i--;
                        }
                        else if (isFieldStart && entity.EntityType == EntityType.FieldMark
                            && (entity as WFieldMark).Type == FieldMarkType.FieldEnd)
                        {
                            paragraph.ChildEntities.Remove(entity);
                            isFieldStart = false;
                            i--;
                        }
                        else if (isFieldStart)
                        {
                            paragraph.ChildEntities.Remove(entity);
                            i--;
                        }
                    }
                }
            }
        }
    }
}
