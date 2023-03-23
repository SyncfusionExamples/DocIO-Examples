using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Media;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Open_and_save_Word_document
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenAndSave_Click(object sender, RoutedEventArgs e)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Automatic))
            {
                //Access the section in a Word document.
                IWSection section = document.Sections[0];
                    //Add a new paragraph to the section.
                    IWParagraph paragraph = section.AddParagraph();
                    paragraph.ParagraphFormat.FirstLineIndent = 36;
                    paragraph.BreakCharacterFormat.FontSize = 12f;
                    IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
                    text.CharacterFormat.FontSize = 12f;

                    //Save the Word document.
                    document.Save(Path.GetFullPath(@".. /../Sample.docx"));               
            }
            MessageBox.Show("Word document generated successfully");
        }
    }
}
