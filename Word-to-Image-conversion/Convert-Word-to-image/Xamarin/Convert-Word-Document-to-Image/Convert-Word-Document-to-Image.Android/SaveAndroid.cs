using System;
using System.IO;
using Convert_Word_Document_to_Image.Droid;
using Android.Content;
using Java.IO;
using Xamarin.Forms;
using System.Threading.Tasks;
using AndroidX.Core.Content;


[assembly: Dependency(typeof(SaveAndroid))]

class SaveAndroid : ISave
{
    //Method to save document as a file in Android and view the saved document
    public async Task SaveAndView(string fileName, String contentType, MemoryStream stream)
    {
        string root = null;
        //Get the root path in android device.
        if(Android.OS.Environment.IsExternalStorageEmulated)
            {
            root = Android.App.Application.Context!.GetExternalFilesDir(Android.OS.Environment.DirectoryDownloads)!.AbsolutePath;
        }
            else
            root = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);

        //Create directory and file 
        Java.IO.File myDir = new Java.IO.File(root + "/Syncfusion");
        myDir.Mkdir();

        Java.IO.File file = new Java.IO.File(myDir, fileName);

        //Remove if the file exists
        if (file.Exists()) file.Delete();

        //Write the stream into the file
        try
        {
            FileOutputStream outs = new FileOutputStream(file);
            stream.Position = 0;
            outs.Write(stream.ToArray());

            outs.Flush();
            outs.Close();
        }
        catch(System.Exception e)
        {
            string exception = e.ToString();
        }
      

        //Invoke the created file for viewing
        if (file.Exists())

        {

            Android.Net.Uri path = Android.Net.Uri.FromFile(file);

            string extension = Android.Webkit.MimeTypeMap.GetFileExtensionFromUrl(Android.Net.Uri.FromFile(file).ToString());

            string mimeType = Android.Webkit.MimeTypeMap.Singleton.GetMimeTypeFromExtension(extension);

            Intent intent = new Intent(Intent.ActionView);

            intent.SetFlags(ActivityFlags.ClearTop | ActivityFlags.NewTask);

            path = FileProvider.GetUriForFile(Android.App.Application.Context, Android.App.Application.Context.PackageName + ".provider", file);

            intent.SetDataAndType(path, mimeType);

            intent.AddFlags(ActivityFlags.GrantReadUriPermission);

            Android.App.Application.Context.StartActivity(intent);

        }
    }
}
