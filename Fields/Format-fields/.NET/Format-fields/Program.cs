using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Format_fields
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of a WordDocument.
            using (WordDocument document = new WordDocument())
            {
                //Adds one section and one paragraph to the document.
                document.EnsureMinimal();
                //Adds the new Page field in Word document with field name and its type.
                IWField field = document.LastParagraph.AppendField("Page", FieldType.FieldPage);
                IEntity entity = field;
                //Iterates to sibling items until Field End.
                while (entity.NextSibling != null)
                {
                    if (entity is WTextRange)
                        //Sets character format for text ranges.
                        (entity as WTextRange).CharacterFormat.FontSize = 6;
                    else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldEnd)
                        break;
                    //Gets next sibling item.
                    entity = entity.NextSibling;
                }
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
