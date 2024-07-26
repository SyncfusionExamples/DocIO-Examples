using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using System.Reflection;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Convert_Word_document_to_Image;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Convert_Word_Document_to_Image
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
        }
        private void ConvertWordtoImage(object sender, RoutedEventArgs e)
        {
            //Loading an existing Word document
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Convert_Word_Document_to_Image.Assets.Input.docx"), FormatType.Docx))
            {
                //Instantiation of DocIORenderer for Word to Image conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Convert the first page of the Word document into an image.
                    Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    imageStream.Position = 0;
                    SaveHelper.SaveAndLaunch("WordToImage.Jpeg", imageStream as MemoryStream);
                }
            }
        }
    }
}
