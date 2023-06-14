using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Imaging;

namespace Convert_Word_Document_to_Image
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnConvert_Click(object sender, EventArgs e)
        {
            //Load the existing Word document 
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Automatic))
            {
                //Convert the first page of the Word document into an image.
                Image image = document.RenderAsImages(0, ImageType.Bitmap);
                //Save the image as jpeg.
                image.Save(Path.GetFullPath(@"../../WordToImage.Jpeg"));
            }

            //Launch the Image file
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../WordToImage.Jpeg")) { UseShellExecute = true };
            System.Diagnostics.Process.Start(Path.GetFullPath(@"../../WordToImage.Jpeg"));
        }
    }
}
