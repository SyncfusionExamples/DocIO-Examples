using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Windows.Storage.Pickers;
using Windows.Storage;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using System.Reflection;
using Syncfusion.DocIO;


// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Convert_Word_Document_to_PDF
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
            //Load an existing Word document.
            Assembly assembly = typeof(App).GetTypeInfo().Assembly;
            using (WordDocument wordDocument = new WordDocument(assembly.GetManifestResourceStream("Convert_Word_Document_to_PDF.Assets.Input.docx"), FormatType.Docx))
            {
                //Instantiation of DocIORenderer for Word to PDF conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Converts Word document into PDF document
                    using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
                    {                       
                        //Saves the PDF file
                        MemoryStream outputStream = new MemoryStream();
                        pdfDocument.Save(outputStream);

                        //Closes the instance of PDF document object
                        pdfDocument.Close();
                        outputStream.Position = 0;

                        //Save the PDF file
                        SavePDF(outputStream);
                    }                   
                }               
            }               
        }


        /// <summary>
        /// Save the PDF to the desired file location
        /// </summary>
        private async void SavePDF(Stream outputStream)
        {
            StorageFile stFile;
            if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
            {
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".pdf";
                savePicker.SuggestedFileName = "Sample";
                savePicker.FileTypeChoices.Add("Adobe PDF Document", new List<string>() { ".pdf" });
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync("Sample.pdf", CreationCollisionOption.ReplaceExisting);
            }
            if (stFile != null)
            {
                Windows.Storage.Streams.IRandomAccessStream fileStream = await stFile.OpenAsync(FileAccessMode.ReadWrite);
                Stream st = fileStream.AsStreamForWrite();
                st.SetLength(0);
                st.Write((outputStream as MemoryStream).ToArray(), 0, (int)outputStream.Length);
                st.Flush();
                st.Dispose();
                fileStream.Dispose();
                MessageDialog msgDialog = new MessageDialog("Do you want to view the Document?", "File created.");
                UICommand yesCmd = new UICommand("Yes");
                msgDialog.Commands.Add(yesCmd);
                UICommand noCmd = new UICommand("No");
                msgDialog.Commands.Add(noCmd);
                IUICommand cmd = await msgDialog.ShowAsync();
                if (cmd == yesCmd)
                {
                    // Launch the retrieved file
                    bool success = await Windows.System.Launcher.LaunchFileAsync(stFile);
                }
            }
        }

    }
}
