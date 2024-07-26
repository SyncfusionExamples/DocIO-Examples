using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Convert_Word_Document_to_Image
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void OnButtonClicked(object sender, EventArgs e)
        {
            //Loading an existing Word document
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;

            using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Convert-Word-Document-to-Image.Assets.Input.docx"), FormatType.Docx))
            {
                //Instantiation of DocIORenderer for Word to Image conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Convert the first page of the Word document into an image.
                    Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    imageStream.Position = 0;
                    //Save the stream as file.
                    Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("WordToImage.Jpeg", "application/jpeg", imageStream as MemoryStream);

                }
            }
        }
    }
}
