using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Mail_merge_date_in_German_culture
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Execute mail merge
                    string[] fieldnames = { "Name", "Date" };
                    string[] fieldvalues = { "Andrew", DateTime.Now.ToString() };

                    //Hook the even to do the date format changes during mail merge
                    document.MailMerge.MergeField += ChangeDateLanguauge;
                    document.MailMerge.Execute(fieldnames, fieldvalues);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        static void ChangeDateLanguauge(object sender, MergeFieldEventArgs args)
        {
            //Check whether date is merge for this merge field
            if (args.FieldName == "Date")
            {
                //Get the date value
                string dateValue = args.FieldValue.ToString();
                //Parse the date value
                DateTime date = DateTime.Parse(dateValue);
                //Convert the date value to German culture in the same date format 
                string formattedDate = date.ToString(args.CurrentMergeField.DateFormat, new System.Globalization.CultureInfo("de-DE"));
                //Set the date value to the current merge field
                args.Text = formattedDate;

            }
        }
    }
}
