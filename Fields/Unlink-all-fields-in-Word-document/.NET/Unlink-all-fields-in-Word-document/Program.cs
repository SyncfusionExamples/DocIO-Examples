using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.Collections.Generic;
using System.IO;

namespace Unlink_all_fields_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Updates the fields present in a document
                    document.UpdateDocumentFields(true);

                    //Get all fields present in a document.
                    List<Entity> fields = document.FindAllItemsByProperty(EntityType.Field, null, null);

                    //Unlink all the fields.
                    for (int i = 0; i < fields.Count; i++)
                    {
                        WField field = (WField)fields[i];
                        if (field.Owner != null)
                            field.Unlink();
                    }

                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
	}
}
