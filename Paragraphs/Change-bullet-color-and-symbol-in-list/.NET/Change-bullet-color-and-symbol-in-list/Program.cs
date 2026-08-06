using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;


//Load the input Word document.
using (WordDocument document = new WordDocument(@"Data/Input.docx", FormatType.Docx))
{
    //Access the list style in a Word document.
    ListStyle style = document.ListStyles[0];
    WListLevel levelOne = style.Levels[0];
    //Define the character and pattern for level 1.
    levelOne.PatternType = ListPatternType.Bullet;
    levelOne.BulletCharacter = "\u0076";
    levelOne.CharacterFormat.FontName = "Wingdings";
    levelOne.CharacterFormat.TextColor = Color.Red;
    //Save the Word document.
    document.Save(@"Output/Sample.docx", FormatType.Docx);
}

