using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace User_defined_numbered_list
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
                //Defines the follow character, prefix, suffix, start index for level 0.
                levelOne.FollowCharacter = FollowCharacterType.Tab;
                levelOne.NumberPrefix = "(";
                levelOne.NumberSuffix = ")";
                levelOne.PatternType = ListPatternType.LowRoman;
                levelOne.StartAt = 1;
                levelOne.TabSpaceAfter = 5;
                levelOne.NumberAlignment = ListNumberAlignment.Center;
                WListLevel levelTwo = listStyle.Levels[1];
                //Defines the follow character, suffix, pattern, start index for level 1.
                levelTwo.FollowCharacter = FollowCharacterType.Tab;
                levelTwo.NumberSuffix = "}";
                levelTwo.PatternType = ListPatternType.LowLetter;
                levelTwo.StartAt = 2;
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
