using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace List_with_prefix_from_previous_level
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new list style to the document.
                ListStyle listStyle = document.AddListStyle(ListType.Numbered, "UserDefinedList");
                WListLevel levelOne = listStyle.Levels[0];
                //Defines the follow character, prefix from previous level, start index for level 0.
                levelOne.FollowCharacter = FollowCharacterType.Nothing;
                levelOne.PatternType = ListPatternType.Arabic;
                levelOne.StartAt = 1;
                WListLevel levelTwo = listStyle.Levels[1];
                //Defines the follow character, prefix from previous level, pattern, start index for level 1.
                levelTwo.FollowCharacter = FollowCharacterType.Nothing;
                levelTwo.NumberPrefix = "\u0000.";
                levelTwo.PatternType = ListPatternType.Arabic;
                levelTwo.StartAt = 1;
                WListLevel levelThree = listStyle.Levels[2];
                //Defines the follow character, prefix from previous level, pattern, start index for level 1.
                levelThree.FollowCharacter = FollowCharacterType.Nothing;
                levelThree.NumberPrefix = "\u0000.\u0001.";
                levelThree.PatternType = ListPatternType.Arabic;
                levelThree.StartAt = 1;
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Adds text to the paragraph.
                paragraph.AppendText("User defined list - Level 0");
                //Applies default numbered list style.
                paragraph.ListFormat.ApplyStyle("UserDefinedList");
                //Adds second paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("User defined list - Level 1");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increases the level indent.
                paragraph.ListFormat.IncreaseIndentLevel();
                //Adds second paragraph.
                paragraph = section.AddParagraph();
                paragraph.AppendText("User defined list - Level 2");
                //Continues last defined list.
                paragraph.ListFormat.ContinueListNumbering();
                //Increases the level indent.
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
