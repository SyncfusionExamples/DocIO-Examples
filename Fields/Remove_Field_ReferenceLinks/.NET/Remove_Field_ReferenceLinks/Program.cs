using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;


namespace Remove_Field_ReferenceLink
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the template Word document.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Get all reference fields present in a document.
                    List<Entity> fields = document.FindAllItemsByProperty(EntityType.Field, "FieldType", FieldType.FieldRef.ToString());
                    //Unlink all ref fields.
                    for (int i = 0; i < fields.Count; i++)
                    {
                        WField field = (WField)fields[i];
                        if (field.Owner != null)
                            field.Unlink();
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
