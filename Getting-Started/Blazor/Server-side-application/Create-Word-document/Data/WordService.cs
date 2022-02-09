using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_Word_document.Data
{
    public class WordService
    {
        public MemoryStream CreateWord()
        {
            //Creating a new document.
            using (WordDocument document = new WordDocument())
            {
                //Adding a new section to the document.
                WSection section = document.AddSection() as WSection;
                //Set Margin of the section.
                section.PageSetup.Margins.All = 72;
                //Set page size of the section.
                section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(612, 792);

                //Create Paragraph styles.
                WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
                style.CharacterFormat.FontName = "Calibri";
                style.CharacterFormat.FontSize = 11f;
                style.ParagraphFormat.BeforeSpacing = 0;
                style.ParagraphFormat.AfterSpacing = 8;
                style.ParagraphFormat.LineSpacing = 13.8f;

                style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
                style.ApplyBaseStyle("Normal");
                style.CharacterFormat.FontName = "Calibri Light";
                style.CharacterFormat.FontSize = 16f;
                style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.FromArgb(46, 116, 181);
                style.ParagraphFormat.BeforeSpacing = 12;
                style.ParagraphFormat.AfterSpacing = 0;
                style.ParagraphFormat.Keep = true;
                style.ParagraphFormat.KeepFollow = true;
                style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
                IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

                paragraph.ApplyStyle("Normal");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                WTextRange textRange = paragraph.AppendText("Adventure Works Cycles") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Calibri";
                textRange.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Red;

                //Appends paragraph.
                paragraph = section.AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                textRange = paragraph.AppendText("Adventure Works Cycles") as WTextRange;
                textRange.CharacterFormat.FontSize = 18f;
                textRange.CharacterFormat.FontName = "Calibri";

                //Appends paragraph.
                paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                textRange = paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company. The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets. While its base operation is in Bothell, Washington with 290 employees, several regional sales teams are located throughout their market base.") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;

                //Appends paragraph.
                paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                textRange = paragraph.AppendText("In 2000, AdventureWorks Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;

                //Saves the Word document to MemoryStream.
                MemoryStream stream = new MemoryStream();
                document.Save(stream, FormatType.Docx);
                stream.Position = 0;
                return stream;
            }
        }
    }
}
