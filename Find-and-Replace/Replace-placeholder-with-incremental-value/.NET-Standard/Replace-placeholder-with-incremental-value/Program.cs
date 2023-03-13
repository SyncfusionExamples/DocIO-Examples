using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_placeholder_with_incremental_value
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2VVhkQlFadV5JXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0djXn5ZcXVQRWVfVEA=");
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Find the occurrence of the Word "{counter}" in the document.
                    TextSelection[] textSelection = document.FindAll("{counter}", false, false);
                    //Iterate through each occurrence and change the text as incremental value.
                    int counter = 1;
                    foreach (TextSelection selection in textSelection)
                    {
                        IWTextRange textRange = selection.GetAsOneRange();
                        textRange.Text = counter.ToString();
                        counter++;
                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }               
            }
        }
    }
}
