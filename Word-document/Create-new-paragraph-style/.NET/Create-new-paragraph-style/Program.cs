using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Create_new_paragraph_style
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    IWParagraphStyle myStyle = document.AddParagraphStyle("MyStyle");
                    //Sets the formatting of the style.
                    myStyle.CharacterFormat.FontSize = 16f;
                    myStyle.CharacterFormat.TextColor = Syncfusion.Drawing.Color.DarkBlue;
                    myStyle.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                    //Appends the contents into the paragraph.
                    document.LastParagraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                    //Applies the style to paragraph.
                    document.LastParagraph.ApplyStyle("MyStyle");
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
