using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.DocIO.DLS;
using System.Reflection;
using Windows.Storage.Pickers;
using Windows.Storage;
using Windows.UI.Popups;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Convert_Word_Document_to_Image
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
            using (WordDocument wordDocument = new WordDocument(assembly.GetManifestResourceStream("Convert_Word_Document_to_Image.Assets.Input.docx"), FormatType.Docx))
            {
                //Instantiation of DocIORenderer for Word to Image conversion
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Convert the first page of the Word document into an image.
                    Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    imageStream.Position = 0;
                    //Save the image file.
                    SaveImage(imageStream);
                }
            }
        }

        /// <summary>
        /// Save the Image to the desired file location
        /// </summary>
        private async void SaveImage(Stream outputStream)
        {
            StorageFile stFile;
            if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))
            {
                FileSavePicker savePicker = new FileSavePicker();
                savePicker.DefaultFileExtension = ".jpeg";
                savePicker.SuggestedFileName = "Sample";
                savePicker.FileTypeChoices.Add("Image viewer ", new List<string>() { ".jpeg" });
                stFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
                stFile = await local.CreateFileAsync("WordToImage.Jpeg", CreationCollisionOption.ReplaceExisting);
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
