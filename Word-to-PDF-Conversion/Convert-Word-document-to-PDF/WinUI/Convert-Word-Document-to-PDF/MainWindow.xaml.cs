using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace Convert_Word_Document_to_PDF
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

        private void ConvertWordtoPDF(object sender, RoutedEventArgs e)
        {
            //Loading an existinG Word document
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            using (WordDocument document = new WordDocument(assembly.GetManifestResourceStream("Convert_Word_Document_to_PDF.Assets.Input.docx"), FormatType.Docx))
            {
                //Instantiation of DocIORenderer for Word to PDF conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Converts Word document into PDF document
                    using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                    {
                        //Save a PDF document to the stream.
                        using MemoryStream outputStream = new();
                        pdfDocument.Save(outputStream);

                        //Save the stream as a Word document file in the local machine.
                        SaveHelper.SaveAndLaunch("Sample.pdf", outputStream);
                    }
                }
            }
        }
    }
}
