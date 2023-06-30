using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Find_and_replace_text_with_image
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnCreate_Click(object sender, EventArgs e)
        {
            //Creating a new document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Docx))
            {
                //Finds  image placeholder text in the Word document
                TextSelection textSelection = document.Find("{ImagePlaceHolder}", false, false);
                //Replaces the image placeholder text with desired image
                WParagraph paragraph = new WParagraph(document);
                WPicture picture = paragraph.AppendPicture(Image.FromFile(Path.GetFullPath(@"../../Data/Image.png"))) as WPicture;
                TextSelection newSelection = new TextSelection(paragraph, 0, 1);
                TextBodyPart bodyPart = new TextBodyPart(document);
                bodyPart.BodyItems.Add(paragraph);
                document.Replace(textSelection.SelectedText, bodyPart, true, true);
                //Saves and closes the document
                document.Save(Path.GetFullPath(@"../../Sample.docx"));
            }
            MessageBox.Show("Word document generated successfully");
        }
    }
}
