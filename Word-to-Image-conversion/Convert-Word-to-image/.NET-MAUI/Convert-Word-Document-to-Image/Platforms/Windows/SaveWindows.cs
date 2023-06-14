using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.Storage;
using Windows.UI.Popups;

namespace Convert_Word_Document_to_Image.SaveServices
{
    public partial class SaveService
    {
        public async partial void SaveAndView(string filename, string contentType, MemoryStream stream)
        {
            StorageFile storageFile;
            string extension = Path.GetExtension(filename);
            //Gets process windows handle to open the dialog in application process.
            IntPtr windowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            if (!Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons"))
            {
                FileSavePicker savePicker = new();
                if (extension == ".Jpeg")
                {
                    savePicker.DefaultFileExtension = ".jpeg";
                    savePicker.SuggestedFileName = filename;
                    //Saves the file as Image file.
                    savePicker.FileTypeChoices.Add("JPEG", new List<string>() { ".jpeg" });
                }

                WinRT.Interop.InitializeWithWindow.Initialize(savePicker, windowHandle);
                storageFile = await savePicker.PickSaveFileAsync();
            }
            else
            {
                StorageFolder local = ApplicationData.Current.LocalFolder;
                storageFile = await local.CreateFileAsync(filename, CreationCollisionOption.ReplaceExisting);
            }
            if (storageFile != null)
            {
                try
                {
                    using (IRandomAccessStream zipStream = await storageFile.OpenAsync(FileAccessMode.ReadWrite))
                    {
                        //Writes compressed data from memory to file.
                        using Stream outstream = zipStream.AsStreamForWrite();
                        outstream.SetLength(0);
                        byte[] buffer = stream.ToArray();
                        outstream.Write(buffer, 0, buffer.Length);
                        outstream.Flush();
                    }
                    //Creates message dialog box. 
                    MessageDialog msgDialog = new("Do you want to view the Document?", "File has been created successfully");
                    UICommand yesCmd = new("Yes");
                    msgDialog.Commands.Add(yesCmd);
                    UICommand noCmd = new("No");
                    msgDialog.Commands.Add(noCmd);

                    WinRT.Interop.InitializeWithWindow.Initialize(msgDialog, windowHandle);

                    //Showing a dialog box. 
                    IUICommand cmd = await msgDialog.ShowAsync();
                    if (cmd.Label == yesCmd.Label)
                    {
                        if (extension == ".md")
                        {
                            Windows.System.LauncherOptions options = new Windows.System.LauncherOptions();
                            options.DisplayApplicationPicker = true;
                            // Launch the file with "Open with" option.
                            await Windows.System.Launcher.LaunchFileAsync(storageFile, options);
                        }
                        else
                            //Launch the saved file. 
                            await Windows.System.Launcher.LaunchFileAsync(storageFile);
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Access is denied."))
                    {
                        //Creates message dialog box.
                        MessageDialog msgDialogBox = new("Access to the given path is denied. Please enable permission to save the file in that folder or save the file in another location.", "Access Denied");
                        UICommand okCmd = new("Ok");
                        msgDialogBox.Commands.Add(okCmd);
                        WinRT.Interop.InitializeWithWindow.Initialize(msgDialogBox, windowHandle);
                        //Showing a dialog box. 
                        IUICommand msgCmd = await msgDialogBox.ShowAsync();
                    }
                }
            }
        }
        

    }
}
    
