using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.IO;
using System.Drawing;

namespace Convert_Word_Document_to_Image
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
        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Docx))
            {
                //Convert the first page of the Word document into an image.
                System.Drawing.Image image = document.RenderAsImages(0, ImageType.Bitmap);
                //Save the image as jpeg.
                image.Save(Path.GetFullPath(@"../../WordToImage.Jpeg"));
            }

            //Launch the  Image
            if (System.Windows.MessageBox.Show("Do you want to view the Image?", "Image created",
  MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../WordToImage.Jpeg")) { UseShellExecute = true };
                process.Start();
            }
        }
    }
}
