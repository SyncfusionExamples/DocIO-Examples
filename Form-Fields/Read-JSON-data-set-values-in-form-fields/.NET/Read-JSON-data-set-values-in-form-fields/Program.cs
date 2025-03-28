using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Read_JSON_data_set_values_in_form_fields
{
    class Program
    {
        static void Main(string[] args)
        {
            // Read and deserialize JSON data.
            var jsonString = File.ReadAllText(Path.GetFullPath("Data/ReportData.json"));
            var data = JsonConvert.DeserializeObject<Root>(jsonString);

            using (WordDocument document = new WordDocument())
            {
                ////Get the text from json and replace in Word document
                //Adds new section to the document
                IWSection section = document.AddSection();
                //Adds new paragraph to the section
                WParagraph paragraph = section.AddParagraph().AppendText("General Information") as WParagraph;
                section.AddParagraph();

                AddTextFormField(section, "EmployeeID", data.Reports[0].EmployeeID);
                AddTextFormField(section, "Name", data.Reports[0].Name);
                AddTextFormField(section, "PhoneNumber", data.Reports[0].PhoneNumber);
                AddTextFormField(section, "Location", data.Reports[0].Location);

                //Creates an instance of DocIORenderer.
                using (DocIORenderer renderer = new DocIORenderer())
                {
                    //Converts Word document into PDF document.
                    using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
                    {
                        //Saves the PDF file to file system.    
                        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            pdfDocument.Save(outputStream);
                        }
                    }
                }
            }
        }

        private static void AddTextFormField(IWSection section, string label, string value)
        {
            WParagraph paragraph = section.AddParagraph() as WParagraph;
            IWTextRange text = paragraph.AppendText(label + ": ");
            text.CharacterFormat.Bold = true;

            WTextFormField textField = paragraph.AppendTextFormField(null);
            textField.Type = TextFormFieldType.RegularText;
            textField.CharacterFormat.FontName = "Calibri";
            textField.Text = value;
            textField.CalculateOnExit = true;

            section.AddParagraph();
        }
    }
    public class Reports
    {
        [JsonProperty("EmployeeID")] public string EmployeeID { get; set; }
        [JsonProperty("Name")] public string Name { get; set; }
        [JsonProperty("PhoneNumber")] public string PhoneNumber { get; set; }
        [JsonProperty("Location")] public string Location { get; set; }

    }

    public class Root
    {
        [JsonProperty("Reports")] public List<Reports> Reports { get; set; }
    }
}
