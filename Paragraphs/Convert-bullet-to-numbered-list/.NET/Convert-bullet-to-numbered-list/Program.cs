using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;

string inputFile = @"../../../Data/Input.docx";
string outputFile = @"../../../Output/Ouput.docx";

WordDocument wordDocument = new WordDocument(inputFile, FormatType.Docx);
// Replace {list-id} with the bullet list's id you want to target.
List<Entity> bullets = wordDocument.FindAllItemsByProperty(EntityType.Paragraph,"ListFormat.ListType","Bulleted");

foreach (Entity entity in bullets)
    (entity as WParagraph).ListFormat.ApplyDefNumberedStyle();

wordDocument.Save(outputFile);
wordDocument.Close();

