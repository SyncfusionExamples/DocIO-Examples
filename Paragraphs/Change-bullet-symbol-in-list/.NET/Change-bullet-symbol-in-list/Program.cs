using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Load the file stream into a Word document.
WordDocument document = new WordDocument(inputStream, FormatType.Docx);
//Access the list style in a Word document.
ListStyle style = document.ListStyles[0];
WListLevel levelOne = style.Levels[0];
//Define the character and pattern for level 1.
levelOne.PatternType = ListPatternType.Bullet;
levelOne.BulletCharacter = "\u0076";
levelOne.CharacterFormat.FontName = "Wingdings";
levelOne.CharacterFormat.TextColor = Color.Red;
//Create a file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Save the Word document to the file stream.
    document.Save(outputFileStream, FormatType.Docx);
}

