using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Change_bullet_symbol_in_list
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Access the list style in a Word document.
                    ListStyle style = document.ListStyles[0];
                    WListLevel levelOne = style.Levels[0];
                    //Define the character and pattern for level 1.
                    levelOne.PatternType = ListPatternType.Bullet;
                    levelOne.BulletCharacter = "\u0076";
                    levelOne.CharacterFormat.FontName = "Wingdings";
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
