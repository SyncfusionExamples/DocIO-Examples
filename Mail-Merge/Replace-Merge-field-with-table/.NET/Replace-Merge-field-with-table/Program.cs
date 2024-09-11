using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;

namespace Replace_Merge_field_with_table
{
    class Program
    {
        static Dictionary<WParagraph, Dictionary<int, WTable>> paraToInsertTable = new Dictionary<WParagraph, Dictionary<int, WTable>>();
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Enables the flag to start each record in new page.
                    document.MailMerge.StartAtNewPage = true;
                    //Gets the employee details as “IEnumerable” collection
                    List<Employees> employeeList = GetEmployeeData(document);
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Employees", employeeList);
                    //Uses the mail merge event handler to insert chart during mail merge.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(MergeField_Table);
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataTable);
                    InsertTable();
                    //Unhooks the event after mail merge execution.
                    document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeField_Table);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"Output/Result.docx")) { UseShellExecute = true });
        }
        #region Helper methods
        /// <summary>
        /// Gets the employee data to perform mail merge. 
        /// </summary>
        /// <returns></returns>
        public static List<Employees> GetEmployeeData(WordDocument document)
        {
            WTable table = CreateTable(document);
            //Adds all details in employee data collection for all employees.
            List<Employees> employeeData = new List<Employees>();
            employeeData.Add(new Employees("Nancy", "Davolio", "1", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "USA", table));
           
            return employeeData;
        }
        /// <summary>
        /// Creates the table.
        /// </summary>
        /// <param name="document"></param>
        private static WTable CreateTable(WordDocument document)
        {
            //Adds a new table into Word document
            WTable table = new WTable(document);
            //Specifies the total number of rows & columns
            table.ResetCells(3, 2);
            //Accesses the instance of the cell (first row, first cell) and adds the content into cell
            IWTextRange textRange = table[0, 0].AddParagraph().AppendText("Item");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (first row, second cell) and adds the content into cell
            textRange = table[0, 1].AddParagraph().AppendText("Number of items sold out");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;
            //Accesses the instance of the cell (second row, first cell) and adds the content into cell
            textRange = table[1, 0].AddParagraph().AppendText("Mountain-350");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;
            //Accesses the instance of the cell (second row, second cell) and adds the content into cell
            textRange = table[1, 1].AddParagraph().AppendText("50");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;
            //Accesses the instance of the cell (third row, first cell) and adds the content into cell
            textRange = table[2, 0].AddParagraph().AppendText("Mountain-500");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;
            //Accesses the instance of the cell (third row, second cell) and adds the content into cell
            textRange = table[2, 1].AddParagraph().AppendText("30");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;
            return table;
        }
        /// <summary>
        /// Represents the method that handles MergeField event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private static void MergeField_Table(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldName == "TableDetails")
            {
                //Gets the current merge field owner paragraph.
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                WTextBody ownerTextBody = paragraph.OwnerTextBody;
                //Gets the index of the owner paragraph.
                int paraIndex = ownerTextBody.ChildEntities.IndexOf(args.CurrentMergeField.OwnerParagraph);
                //Maintain table in collection.
                Dictionary<int, WTable> fieldValues = new Dictionary<int, WTable>();
                fieldValues.Add(paraIndex, args.FieldValue as WTable);
                //Maintain paragraph in collection.
                paraToInsertTable.Add(paragraph, fieldValues);
                //Set field value as empty.
                args.Text = string.Empty;
            }
        }
        /// <summary>
        /// Append Table to Textbody.
        /// </summary>
        private static void InsertTable()
        {
            //Iterates through each item in the dictionary.
            foreach (KeyValuePair<WParagraph, Dictionary<int, WTable>> dictionaryItems in paraToInsertTable)
            {
                WParagraph paragraph = dictionaryItems.Key;
                Dictionary<int, WTable> values = dictionaryItems.Value;
                //Iterates through each value in the dictionary.
                foreach (KeyValuePair<int, WTable> valuePair in values)
                {

                    int index = valuePair.Key;
                    WTable fieldValue = valuePair.Value;
                    //Inserts table at the same position of mergefield in Word document.
                    paragraph.OwnerTextBody.ChildEntities.Insert(index, fieldValue);
                }
            }
            paraToInsertTable.Clear();
        }
        #endregion
    }
    #region Helper Class
    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Employees
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmployeeID { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public WTable TableDetails { get; set; }
        public Employees(string firstName, string lastName, string employeeID, string address, string city, string country, WTable tableDetails)
        {
            FirstName = firstName;
            LastName = lastName;
            EmployeeID = employeeID;
            Address = address;
            City = city;
            Country = country;
            TableDetails = tableDetails;
        }
    }
    #endregion
}