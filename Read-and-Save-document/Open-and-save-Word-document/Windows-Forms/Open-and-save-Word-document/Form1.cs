using System;
using System.IO;
using System.Windows.Forms;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Open_and_save_Word_document
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpenAndSave_Click(object sender, EventArgs e)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Automatic))
            {
                //Access the section in a Word document.
                IWSection section = document.Sections[0];
                //Add new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                //Add new text to the paragraph.
                WTextRange textRange = paragraph.AppendText("In 2000, AdventureWorks Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.") as WTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                //Save the Word document.
                document.Save(Path.GetFullPath(@"../../Sample.docx"), FormatType.Docx);
            }
            MessageBox.Show("Word document generated successfully");
        }
    }
}
