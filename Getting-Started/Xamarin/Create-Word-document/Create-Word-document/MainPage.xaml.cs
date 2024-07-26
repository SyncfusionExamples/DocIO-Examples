using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Create_Word_document
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void OnButtonClicked(object sender, EventArgs e)
        {
            //"App" is the class of Portable project.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;

            //Creating a new document
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

                //Gets the image stream.
                Stream imageStream1 = assembly.GetManifestResourceStream("Create-Word-document.Templates.AdventureCycle.jpg");
                IWPicture picture = paragraph.AppendPicture(imageStream1);
                picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                picture.VerticalOrigin = VerticalOrigin.Margin;
                picture.VerticalPosition = -45;
                picture.HorizontalOrigin = HorizontalOrigin.Column;
                picture.HorizontalPosition = 263.5f;
                picture.WidthScale = 20;
                picture.HeightScale = 15;

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

                paragraph = section.AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                textRange = paragraph.AppendText("Product Overview") as WTextRange;
                textRange.CharacterFormat.FontSize = 16f;
                textRange.CharacterFormat.FontName = "Calibri";
                //Appends table.
                IWTable table = section.AddTable();
                table.ResetCells(3, 2);
                table.TableFormat.Borders.BorderType = BorderStyle.None;
                table.TableFormat.IsAutoResized = true;

                //Appends paragraph.
                paragraph = table[0, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                //Appends picture to the paragraph.
                Stream imageStream2 = assembly.GetManifestResourceStream("Create-Word-document.Templates.Mountain-200.jpg");
                picture = paragraph.AppendPicture(imageStream2);
                picture.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
                picture.VerticalOrigin = VerticalOrigin.Paragraph;
                picture.VerticalPosition = 4.5f;
                picture.HorizontalOrigin = HorizontalOrigin.Column;
                picture.HorizontalPosition = -2.15f;
                picture.WidthScale = 79;
                picture.HeightScale = 79;

                //Appends paragraph.
                paragraph = table[0, 1].AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.AppendText("Mountain-200");
                //Appends paragraph.
                paragraph = table[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";

                textRange = paragraph.AppendText("Product No: BK-M68B-38\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Size: 38\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Weight: 25\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Price: $2,294.99\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                //Appends paragraph.
                paragraph = table[0, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.BreakCharacterFormat.FontSize = 12f;

                //Appends paragraph.
                paragraph = table[1, 0].AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.AppendText("Mountain-300 ");
                //Appends paragraph.
                paragraph = table[1, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Product No: BK-M47B-38\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Size: 35\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Weight: 22\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Price: $1,079.99\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                //Appends paragraph.
                paragraph = table[1, 0].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.BreakCharacterFormat.FontSize = 12f;

                //Appends paragraph.
                paragraph = table[1, 1].AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.LineSpacing = 12f;
                //Appends picture to the paragraph.
                Stream imageStream3 = assembly.GetManifestResourceStream("Create-Word-document.Templates.Mountain-300.jpg");
                picture = paragraph.AppendPicture(imageStream3);
                picture.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
                picture.VerticalOrigin = VerticalOrigin.Paragraph;
                picture.VerticalPosition = 8.2f;
                picture.HorizontalOrigin = HorizontalOrigin.Column;
                picture.HorizontalPosition = -14.95f;
                picture.WidthScale = 75;
                picture.HeightScale = 75;

                //Appends paragraph.
                paragraph = table[2, 0].AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.LineSpacing = 12f;
                //Appends picture to the paragraph.
                Stream imageStream4 = assembly.GetManifestResourceStream("Create-Word-document.Templates.Road-550-W.jpg");
                picture = paragraph.AppendPicture(imageStream4);
                picture.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
                picture.VerticalOrigin = VerticalOrigin.Paragraph;
                picture.VerticalPosition = 3.75f;
                picture.HorizontalOrigin = HorizontalOrigin.Column;
                picture.HorizontalPosition = -5f;
                picture.WidthScale = 92;
                picture.HeightScale = 92;

                //Appends paragraph.
                paragraph = table[2, 1].AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.AppendText("Road-150 ");
                //Appends paragraph.
                paragraph = table[2, 1].AddParagraph();
                paragraph.ParagraphFormat.AfterSpacing = 0;
                paragraph.ParagraphFormat.LineSpacing = 12f;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                paragraph.BreakCharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Product No: BK-R93R-44\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Size: 44\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Weight: 14\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                textRange = paragraph.AppendText("Price: $3,578.27\r") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Times New Roman";
                //Appends paragraph.
                section.AddParagraph();

                //Saves the Word document to MemoryStream.
                MemoryStream stream = new MemoryStream();
                document.Save(stream, FormatType.Docx);

                //Save the stream as a file in the device and invoke it for viewing.
                Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("Sample.docx", "application/msword", stream);
            }
        }
    }
}
