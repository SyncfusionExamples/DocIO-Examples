using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace User_defined_bulleted_list
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a new section to the document.
                IWSection section = document.AddSection();
                //Add a new list style to the document.
                ListStyle listStyle = document.AddListStyle(ListType.Bulleted, "UserDefinedList");
                WListLevel levelOne = listStyle.Levels[0];
                //Define the following character, pattern and start index for level 0.
                levelOne.PatternType = ListPatternType.Bullet;
                levelOne.BulletCharacter = "*";
                levelOne.StartAt = 1;
                WListLevel levelTwo = listStyle.Levels[1];
                //Define the following character, pattern and start index for level 1.
                levelTwo.PatternType = ListPatternType.Bullet;
                levelTwo.BulletCharacter = "\u00A9";
                levelTwo.CharacterFormat.FontName = "Wingdings";
                levelTwo.StartAt = 1;
                WListLevel levelThree = listStyle.Levels[2];
                //Define the following character, pattern and start index for level 2.
                levelThree.PatternType = ListPatternType.Bullet;
                levelThree.BulletCharacter = "\u0076";
                levelThree.CharacterFormat.FontName = "Wingdings";
                levelThree.StartAt = 1;
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Add a text to the paragraph.
                paragraph.AppendText("User defined list - Level 0");
                //Apply the default bulleted list style.
                paragraph.ListFormat.ApplyStyle("UserDefinedList");
                //Add second paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("User defined list - Level 1");
                //Continue the last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increase the level indent.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Add second paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("User defined list - Level 2");
                //Continue the last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increase the level indent.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
