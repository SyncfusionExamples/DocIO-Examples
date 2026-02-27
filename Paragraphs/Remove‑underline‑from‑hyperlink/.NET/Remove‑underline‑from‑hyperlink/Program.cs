using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;


namespace Remove_underline_from_hyperlink
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load an existing Word document into DocIO instance
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Find all hyperlink fields in the document
                    List<Entity> entities = document.FindAllItemsByProperty(EntityType.Field,"FieldType", "FieldHyperlink") as List<Entity>;
                    if (entities != null)
                    {
                        // Process each hyperlink field
                        foreach (IEntity hyperlinkField in entities)
                        {
                            IEntity currentEntity = hyperlinkField;
                            // Iterate through sibling items until reaching the field end
                            while (currentEntity.NextSibling != null)
                            {
                                currentEntity = currentEntity.NextSibling;
                                // Remove underline from text ranges
                                if (currentEntity is WTextRange textRange)
                                {
                                    textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.None;
                                }
                                // Stop at field end marker
                                else if (currentEntity is WFieldMark fieldMark && fieldMark.Type == FieldMarkType.FieldEnd)
                                    break;
                            }
                        }
                    }
                    //Create file stream
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}


