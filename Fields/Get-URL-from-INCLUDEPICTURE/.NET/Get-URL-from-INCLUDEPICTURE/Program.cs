using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System;
using System.IO;
using System.Linq;

namespace Get_URL_from_INCLUDEPICTURE
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Find all INCLUDEPICTURE fields in the Word document based on field type.
                    List<Entity> IncludePictureFields = document.FindAllItemsByProperty(EntityType.Field, "FieldType", "FieldIncludePicture");
                    //List to store extracted URLs.
                    List<string> urls = new List<string>();
                    //Iterate through all INCLUDEPICTURE fields found in the document.
                    foreach (WField field in IncludePictureFields)
                    {
                        //Extract the field code.
                        string fieldCode = field.FieldCode;
                        //The field code is in the format: INCLUDEPICTURE "URL"
                        //Extract the URL between the quotes.
                        string url = ExtractUrlFromFieldCode(fieldCode);
                        //Add the URL to the list.
                        urls.Add(url);
                    }
                    //Print all URLs at the end.
                    Console.WriteLine("INCLUDEPICTURE URLs:");
                    foreach (string url in urls)
                    {
                        Console.WriteLine(url);
                    }
                }
            }
        }
        #region Helper methods
        /// <summary>
        /// Extract the URL from the given INCLUDEPICTURE field code.
        /// </summary>
        /// <param name="fieldCode">The field code containing the URL.</param>
        /// <returns>The extracted URL as a string.</returns>
        static string ExtractUrlFromFieldCode(string fieldCode)
        {
            string url = string.Empty;
            //Find the starting index of the URL (after the first quote).
            int startIndex = fieldCode.IndexOf('"') + 1;
            //Find the ending index of the URL (before the last quote).
            int endIndex = fieldCode.LastIndexOf('"');
            //Ensure valid indices and extract the URL if valid.
            if (startIndex > 0 && endIndex > startIndex)
            {
                url = fieldCode.Substring(startIndex, endIndex - startIndex);
            }
            return url;
        }
        #endregion
    }
}
