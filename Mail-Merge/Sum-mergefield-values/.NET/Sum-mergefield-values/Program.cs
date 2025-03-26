using Newtonsoft.Json.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Sum_mergefield_values
{
     class Program
    {
        static int totalMarks = 0;
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Gets JSON object from JSON string.
                    JObject jsonObject = JObject.Parse(File.ReadAllText(Path.GetFullPath(@"Data/ReportData.json")));
                    // Converts to IDictionary data from JSON object.
                    IDictionary<string, object> data = GetData(jsonObject);

                    //Creates the mail merge data table in order to perform mail merge
                    MailMergeDataTable dataTable = new MailMergeDataTable("Reports", (List<object>)data["Reports"]);

                    //Uses the mail merge event handler to sum the field's values and set that value to the TotalMarks field during mail merge.
                    wordDocument.MailMerge.MergeField += new MergeFieldEventHandler(MergeField_Event);

                    //Performs the mail merge operation with the dynamic collection
                    wordDocument.MailMerge.ExecuteGroup(dataTable);

                    string[] fieldNames = new string[] { "TotalMarks" };
                    string[] fieldValues = new string[] { "" };

                    //Performs the mail merge
                    wordDocument.MailMerge.Execute(fieldNames, fieldValues);

                    //Saves the WOrd document file to file system.    
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                    {
                        wordDocument.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Event to sum the field's values and set that value to the TotalMarks field during mail merge.
        /// </summary>
        private static void MergeField_Event(object sender, MergeFieldEventArgs args)
        {
            
            if (args.FieldName == "Marks")
            {
                //sum the Marks field values.
                totalMarks += Convert.ToInt32(args.FieldValue);
            }
            if (args.FieldName == "TotalMarks")
            {
                //Set sum of the Marks field values to the TotalMarks field;
                args.Text = totalMarks.ToString();
            }
        }

        /// <summary>
        /// Gets array of items from JSON array.
        /// </summary>
        /// <param name="jArray">JSON array.</param>
        /// <returns>List of objects.</returns>
        private static List<object> GetData(JArray jArray)
        {
            List<object> jArrayItems = new List<object>();
            foreach (var item in jArray)
            {
                object keyValue = null;
                if (item is JObject)
                    keyValue = GetData((JObject)item);
                jArrayItems.Add(keyValue);
            }
            return jArrayItems;
        }
        /// <summary>
        /// Gets data from JSON object.
        /// </summary>
        /// <param name="jsonObject">JSON object.</param>
        /// <returns>IDictionary data.</returns>
        private static IDictionary<string, object> GetData(JObject jsonObject)
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();
            foreach (var item in jsonObject)
            {
                object keyValue = null;
                if (item.Value is JArray)
                    keyValue = GetData((JArray)item.Value);
                else if (item.Value is JToken)
                    keyValue = ((JToken)item.Value).ToObject<string>();
                dictionary.Add(item.Key, keyValue);
            }
            return dictionary;
        }
    }
}
