using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


namespace Apply_Styles_For_Inserted_Tables
{
    class Program
    {
        static void Main(string[] args)
        {
            // Opens an input Word template
            using (WordDocument resultDocument = new WordDocument(Path.GetFullPath(@"../../../Data/Template.docx")))
            {
                // Read HTML string from the file.
                string html = File.ReadAllText(Path.GetFullPath(@"../../../Data/Table.html"));

                // Insert HTML (table style is not applied automatically)
                resultDocument.LastSection.Body.InsertXHTML(html);

                // Append table manually - this one applies the table style from the template
                var employees = new List<(string Id, string Name, string Department)>
                {
                    ("1001", "John", "Sales"),
                    ("1002", "Mary", "HR"),
                    ("1003", "David", "IT")
                };

                IWTable table = resultDocument.LastSection.AddTable();
                table.ResetCells(employees.Count + 1, 3);

                // Header row
                table[0, 0].AddParagraph().AppendText("ID");
                table[0, 1].AddParagraph().AppendText("Name");
                table[0, 2].AddParagraph().AppendText("Department");

                // Data rows
                for (int i = 0; i < employees.Count; i++)
                {
                    table[i + 1, 0].AddParagraph().AppendText(employees[i].Id);
                    table[i + 1, 1].AddParagraph().AppendText(employees[i].Name);
                    table[i + 1, 2].AddParagraph().AppendText(employees[i].Department);
                }

                //Finds all the table in the Word document
                List<Entity> tableList = resultDocument.FindAllItemsByProperty(EntityType.Table, "EntityType", EntityType.Table.ToString());
                foreach (var item in tableList)
                {
                    WTable tableInDocument = item as WTable;
                    //Apply table style "TableGrid" to the table
                    tableInDocument.ApplyStyle(BuiltinTableStyle.TableGrid);
                }
                // Save the document to output file
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create))
                {
                    resultDocument.Save(outputStream, FormatType.Docx);
                }
            }
        } 
    }
}

