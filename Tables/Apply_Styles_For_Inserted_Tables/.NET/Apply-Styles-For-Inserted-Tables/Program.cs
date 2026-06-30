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
                IWTable table = resultDocument.LastSection.AddTable();
                //Specify the total number of rows and columns. 
                table.ResetCells(2, 6);

                for (int rowIndex = 0; rowIndex < 2; rowIndex++)
                    for (int columnIndex = 0; columnIndex < 6; columnIndex++)
                        table.Rows[rowIndex].Cells[columnIndex].AddParagraph().Text = (rowIndex * columnIndex).ToString();

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

