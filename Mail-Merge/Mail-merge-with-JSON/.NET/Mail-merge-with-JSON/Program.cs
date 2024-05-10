using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;

namespace Mail_merge_with_JSON
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    // Read JSON data from file.
                    JObject jsonData = JObject.Parse(File.ReadAllText(Path.GetFullPath(@"../../../Json Data.json")));

                    // Prepare data source for mail merge.
                    List<Dictionary<string, object>> dataSource = PrepareDataSource(jsonData);

                    // Perform mail merge with the prepared data source.
                    document.MailMerge.Execute(dataSource);

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
        /// Prepares the data source from JSON data for processing.
        /// </summary>
        private static List<Dictionary<string, object>> PrepareDataSource(JObject jsonData)
        {
            List<Dictionary<string, object>> dataSource = new List<Dictionary<string, object>>();

            // Assuming your JSON structure has an array of employee detail under a key like "Employee".
            JArray employeeArray = jsonData["Employee"] as JArray;

            foreach (JObject employee in employeeArray)
            {
                Dictionary<string, object> studentData = new Dictionary<string, object>();

                // Add employee information to the dictionary.
                foreach (var property in employee.Properties())
                {
                    studentData.Add(property.Name, property.Value);
                }

                dataSource.Add(studentData);
            }
            return dataSource;
        }
    }
}
