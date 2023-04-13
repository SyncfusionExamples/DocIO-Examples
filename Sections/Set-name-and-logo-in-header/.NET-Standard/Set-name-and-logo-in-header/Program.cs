using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_name_and_logo_in_header
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Word document.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Get the Word document section.
                    IWSection section = document.Sections[0];
                    //Add paragraph to the header.
                    IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
                    // Get the image stream.
                    FileStream imageStream = new FileStream(@"../../../Data/Logo.jpg", FileMode.Open, FileAccess.Read);
                    //Append picture to the created paragraph.
                    IWPicture picture = paragraph.AppendPicture(imageStream);
                    //Set the picture properties.
                    picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                    picture.VerticalOrigin = VerticalOrigin.Margin;
                    picture.VerticalPosition = -45;
                    picture.HorizontalOrigin = HorizontalOrigin.Column;
                    picture.HorizontalPosition = 263.5f;
                    picture.WidthScale = 20;
                    picture.HeightScale = 15;

                    //Append the text to the created paragraph.                   
                    WTextRange textRange = paragraph.AppendText("Adventure Works Cycles") as WTextRange;
                    //Apply formatting for the text range.
                    paragraph.ApplyStyle("Normal");
                    textRange.CharacterFormat.FontSize = 12f;
                    textRange.CharacterFormat.FontName = "Calibri";
                    textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Red;
                    paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;

                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
