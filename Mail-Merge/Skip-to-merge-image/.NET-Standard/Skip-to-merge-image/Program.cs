using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Skip_to_merge_image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Uses the mail merge events to perform the conditional formatting during runtime.
                    document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeEmployeePhoto);
                    //Executes Mail Merge with groups.
                    string[] fieldNames = { "Nancy", "Andrew", "Steven" };
                    string[] fieldValues = { Path.GetFullPath(@"../../../Data/Nancy.png"), Path.GetFullPath(@"../../../Data/Andrew.png"), Path.GetFullPath(@"../../../Data/Steven.png") };
                    //Execute mail merge.
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private static void MergeEmployeePhoto(object sender, MergeImageFieldEventArgs args)
        {
            //Skip to merge particular image.
            if (args.FieldName == "Andrew")
                args.Skip = true;
            //Sets image.
            string ProductFileName = args.FieldValue.ToString();
            FileStream imageStream = new FileStream(ProductFileName, FileMode.Open, FileAccess.Read);
            args.ImageStream = imageStream;
            WPicture picture = args.Picture;
            picture.Height = 100;
            picture.Width = 100;
        }
    }
}
