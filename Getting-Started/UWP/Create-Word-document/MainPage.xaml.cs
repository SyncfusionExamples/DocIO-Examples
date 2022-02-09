using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Create_Word_document
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }
        private async void OnButtonClicked(object sender, RoutedEventArgs e)
        {
            //Creating a new document.
            using (WordDocument document = new WordDocument())
            {
                //Adding a new section to the document.
                WSection section = document.AddSection() as WSection;
                //Set Margin of the section.
                section.PageSetup.Margins.All = 72;
                //Set page size of the section.
                section.PageSetup.PageSize = new SizeF(612, 792);

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
                style.CharacterFormat.TextColor = Color.FromArgb(46, 116, 181);
                style.ParagraphFormat.BeforeSpacing = 12;
                style.ParagraphFormat.AfterSpacing = 0;
                style.ParagraphFormat.Keep = true;
                style.ParagraphFormat.KeepFollow = true;
                style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
                IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

                //Gets the image stream.
                Stream imageStream = typeof(MainPage).GetTypeInfo().Assembly.GetManifestResourceStream("Create_Word_document.Assets.AdventureCycle.jpg");
                IWPicture picture = paragraph.AppendPicture(imageStream);
                picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                picture.VerticalOrigin = VerticalOrigin.Margin;
                picture.VerticalPosition = -45;
                picture.HorizontalOrigin = HorizontalOrigin.Column;
                picture.HorizontalPosition = 263.5f;
                picture.WidthScale = 20;
                picture.HeightScale = 15;

                paragraph.ApplyStyle("Normal");
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Left;
                WTextRange textRange = paragraph.AppendText("Adventure Works Cycles") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                textRange.CharacterFormat.FontName = "Calibri";
                textRange.CharacterFormat.TextColor = Color.Red;

                //Appends paragraph.
                paragraph = section.AddParagraph();
                paragraph.ApplyStyle("Heading 1");
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
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
                paragraph.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Left;
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
                imageStream = typeof(MainPage).GetTypeInfo().Assembly.GetManifestResourceStream("Create_Word_document.Assets.Mountain-200.jpg");
                picture = paragraph.AppendPicture(imageStream);
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
                //Appends paragraph
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
                imageStream = typeof(MainPage).GetTypeInfo().Assembly.GetManifestResourceStream("Create_Word_document.Assets.Mountain-300.jpg");
                picture = paragraph.AppendPicture(imageStream);
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
                imageStream = typeof(MainPage).GetTypeInfo().Assembly.GetManifestResourceStream("Create_Word_document.Assets.Road-550-W.jpg");
                picture = paragraph.AppendPicture(imageStream);
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

                MemoryStream stream = new MemoryStream();

                //Saves the Word document to MemoryStream.
                await document.SaveAsync(stream, FormatType.Docx);

                //Saves the stream as Word document file in local machine.
                Save(stream, "Sample.docx");
            }
        }

        //Saves the Word document.
        async void Save(MemoryStream streams, string filename)
        {
            streams.Position = 0;
            StorageFile stFile;
            if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
            {
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".docx";
                savePicker.SuggestedFileName = filename;
                savePicker.FileTypeChoices.Add("Word Documents", new List<string>() { ".docx" });
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync(filename, CreationCollisionOption.ReplaceExisting);
            }
            if (stFile != null)
            {
                using (IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))
                {
                    //Write compressed data from memory to file.
                    using (Stream outstream = zipStream.AsStreamForWrite())
                    {
                        byte[] buffer = streams.ToArray();
                        outstream.Write(buffer, 0, buffer.Length);
                        outstream.Flush();
                    }
                }
            }
            //Launch the saved Word file.
            await Windows.System.Launcher.LaunchFileAsync(stFile);
        }
    }
}
