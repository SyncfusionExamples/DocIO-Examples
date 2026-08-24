using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

string inputFile = @"../../../Input.docx";
string outputFile = @"../../../Output.docx";

WordDocument wordDocument = new WordDocument(inputFile, FormatType.Docx);
// Replace {list-id} with the bullet list's id you want to target.
List<Entity> bullets = wordDocument.FindAllItemsByProperty(EntityType.Paragraph,"ListFormat.ListType","Bulleted");

foreach (Entity entity in bullets)
    (entity as WParagraph).ListFormat.ApplyDefNumberedStyle();

wordDocument.Save(outputFile);
wordDocument.Close();

