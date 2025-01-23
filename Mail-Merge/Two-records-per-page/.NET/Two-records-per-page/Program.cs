using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Data;
using Newtonsoft.Json.Linq;
using System;
using System.IO;

namespace Two_records_per_page
{
    internal class Program
    {
        static int count = 0; // Counter for page breaks.
        static WParagraph endPara; // Variable to store the end paragraph.
        static int index; // Index for the beginning paragraph.

        static void Main(string[] args)
        {
            // Opens the template document
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    // Get the beginning paragraph where group starts.
                    TextSelection selection = document.Find("BeginGroup:Employees", true, true);
                    WParagraph beginPara = selection.GetAsOneRange().OwnerParagraph;

                    // Get the end paragraph where group ends.
                    TextSelection endSelection = document.Find("EndGroup:Employees", true, true);
                    endPara = endSelection.GetAsOneRange().OwnerParagraph.NextSibling as WParagraph;

                    // Start each record on a new page during mail merge.
                    document.MailMerge.StartAtNewPage = true;

                    // Execute the group mail merge
                    document.MailMerge.ExecuteGroup(GetDataTable());

                    // Get index of the beginning paragraph for iteration.
                    index = beginPara.OwnerTextBody.ChildEntities.IndexOf(beginPara);

                    // Remove odd page breaks in document.
                    IterateTextBody(beginPara.OwnerTextBody);

                    // Save the result to a file.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        // Method to iterate through the text body and handle elements.
        static void IterateTextBody(WTextBody textBody)
        {
            // Loop through each entity in the text body (paragraphs, tables, etc.).
            for (int i = index; i < textBody.ChildEntities.Count; i++)
            {
                IEntity bodyItemEntity = textBody.ChildEntities[i];

                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        // Process paragraph unless it's the end paragraph.
                        if (paragraph != endPara)
                            IterateParagraph(paragraph.Items, paragraph);
                        break;
                    case EntityType.Table:
                        // Process tables.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        // Iterate through body items in block content control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }

        // Method to iterate through the table rows and cells.
        static void IterateTable(WTable table)
        {
            // Loop through rows and cells in a table.
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    // Reuse text body iteration logic for table cells.
                    IterateTextBody(cell);
                }
            }
        }

        // Method to iterate through paragraphs and remove page breaks.
        static void IterateParagraph(ParagraphItemCollection paraItems, WParagraph paragraph)
        {
            if (paragraph != null)
            {
                // Loop through the paragraph items in reverse order to check for page breaks.
                for (int i = paraItems.Count - 1; i >= 0; i--)
                {
                    ParagraphItem item = paraItems[i];
                    if (item is Break && (item as Break).BreakType == BreakType.PageBreak)
                    {
                        count++; // Count the page break.
                        if (count % 2 != 0)
                            paraItems.Remove(item); // Remove the page break if count is odd.
                    }
                }
            }
        }

        // Method to get data from a JSON file and convert it to a DataTable for mail merge.
        private static DataTable GetDataTable()
        {
            string jsonString = File.ReadAllText(Path.GetFullPath(@"../../../Data/Data.json"));
            DataTable dataTable = new DataTable("Employees");
            JObject json = JObject.Parse(jsonString);

            bool columnsAdded = false;
            foreach (var item in json["Employees"])
            {
                // Add columns once.
                if (!columnsAdded)
                {
                    foreach (JProperty property in item)
                    {
                        dataTable.Columns.Add(property.Name, typeof(string));
                    }
                    columnsAdded = true;
                }

                // Add rows to DataTable.
                DataRow row = dataTable.NewRow();
                foreach (JProperty property in item)
                {
                    row[property.Name] = property.Value.ToString();
                }
                dataTable.Rows.Add(row);
            }

            return dataTable;
        }
    }
}
