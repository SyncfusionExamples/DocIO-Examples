using Newtonsoft.Json.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace NestedGroup_MailMerge_with_JSON
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Get the Json data details as DataTable
                    List<object> jsonData = GetJsonData();
                    //Creates the mail merge data table in order to perform mail merge
                    MailMergeDataTable dataTable = new MailMergeDataTable("Organizations", jsonData);
                    // Perform mail merge with the prepared data table.
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    //Create file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Prepares the data table from JSON data for processing.
        /// </summary>
        private static List<object> GetJsonData()
        {
            //Reads the JSON object from JSON file.
            JObject jsonObject = JObject.Parse(File.ReadAllText(Path.GetFullPath(@"Data/Data.json")));
            //Converts JSON object to Dictionary.


            IDictionary<string, object> data = GetData(jsonObject);
            return data["Organizations"] as List<object>;
        }

        /// <summary>
        /// Gets data from JSON object.
        /// </summary>
        /// <param name="jsonObject">JSON object.</param>
        /// <returns>Dictionary of data.</returns>
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
        #endregion
    }
}
