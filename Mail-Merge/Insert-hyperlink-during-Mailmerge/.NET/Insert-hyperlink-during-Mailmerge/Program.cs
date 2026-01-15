using Syncfusion.DocIO.DLS;

namespace Insert_hyperlink_during_mailmerge
{
    class Program
    {
        static Dictionary<WParagraph, List<int>> paraToInsertHyperlink = new Dictionary<WParagraph, List<int>>();
        public static void Main(string[] args)
        {
            // Open the template Word document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/EmployeesTemplate.docx"));
            // Gets the employee details as IEnumerable collection.
            List<Employee> employeeList = GetEmployees();
            // Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
            MailMergeDataTable dataSource = new MailMergeDataTable("Employees", employeeList);
            // Uses the mail merge events handler for image fields.
            document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_EmployeeImage);
            // Uses the mail merge event handler for merge fields.
            document.MailMerge.MergeField += MailMerge_MergeField;
            // Performs Mail merge.
            document.MailMerge.ExecuteGroup(dataSource);
            // Insert Hyperlink to the merge field text.
            InsertHyperlink(document);
            // Save the result document
            document.Save(Path.GetFullPath(@"../../../Output/output.docx"));
            // Close the Word document
            document.Close();
        }
        /// <summary>
        /// Event handler that customizes how merge fields are processed during mail merge.
        /// </summary>
        /// <param name="sender">The source object raising the event (MailMerge engine).</param>
        /// <param name="args">Provides details about the current merge field being processed</param>
        private static void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {

            // Check if the current merge field is "Contact"
            if (args.FieldName == "Contact")
            {
                // Get the mergefield's Owner paragraph
                WParagraph mergeFieldOwnerParagraph = args.CurrentMergeField.OwnerParagraph;
                // Check if this paragraph already has an entry in the dictionary.
                // If not, create a new list to store field index.
                if (!paraToInsertHyperlink.TryGetValue(mergeFieldOwnerParagraph, out var fields))
                {
                    fields = new List<int>();
                    paraToInsertHyperlink[mergeFieldOwnerParagraph] = fields;                  
                }
                // Add the current merge field's index
                fields.Add(mergeFieldOwnerParagraph.ChildEntities.IndexOf(args.CurrentMergeField));
            }
        }
        /// <summary>
        /// Inserts hyperlinks into the Word document at the positions of merge fields
        /// </summary>
        /// <param name="document">The WordDocument object being processed.</param>
        private static void InsertHyperlink(WordDocument document)
        {
            foreach (KeyValuePair<WParagraph, List<int>> dictionaryItems in paraToInsertHyperlink)
            {
                // Get the paragraph where Hyperlink needs to be inserted.
                WParagraph paragraph = dictionaryItems.Key;
                // Get the list of index for the merge fields.
                List<int> values = dictionaryItems.Value;
                // Iterate through the list in reverse order
                for (int i = values.Count - 1; i >= 0; i--)
                {
                    // Get the index of the merge field within the paragraph.
                    int index = values[i];
                    // Get the merge field content to insert as Hyperlink.
                    WTextRange mergeFieldText = paragraph.ChildEntities[index] as WTextRange;                                       
                    if (mergeFieldText != null)
                    {
                        string hyperlinkText = mergeFieldText.Text;
                        WParagraph hyperlinkParagraph = new WParagraph(document);
                        WField hyperlink = hyperlinkParagraph.AppendHyperlink(hyperlinkText, hyperlinkText, HyperlinkType.WebLink) as WField;
                        // Insert the child entity (e.g., hyperlink) from the new paragraph into the original paragraph
                        for (int j = hyperlinkParagraph.ChildEntities.Count - 1; j >= 0; j--)
                        {
                            paragraph.ChildEntities.Insert(index, hyperlinkParagraph.ChildEntities[j].Clone());
                        }
                        // Remove the original merge field text from the paragraph
                        paragraph.ChildEntities.Remove(mergeFieldText);
                    }
                }
            }
            paraToInsertHyperlink.Clear();
        }
        /// <summary>
        /// Represents the method that handles MergeImageField event.
        /// </summary>
        private static void MergeField_EmployeeImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge.
            if (args.FieldName == "Photo")
            {
                string photoFileName = args.FieldValue.ToString();
                //Gets the image from file system.
                FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/" + photoFileName), FileMode.Open, FileAccess.Read);
                args.ImageStream = imageStream;
            }
        }
        /// <summary>
        /// Gets the employee details to perform mail merge.
        /// </summary>
        public static List<Employee> GetEmployees()
        {
            List<Employee> employees = new List<Employee>();
            employees.Add(new Employee("Nancy", "Smith", "Sales Representative", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "WA", "USA","nancy.smith@xyz.com", "Nancy.png"));
            employees.Add(new Employee("Andrew", "Fuller", "Vice President, Sales", "908 W. Capital Way", "Tacoma", "WA", "USA", "andrew.fuller@xyz.com", "Andrew.png"));
            employees.Add(new Employee("Roland", "Mendel", "Sales Representative", "722 Moss Bay Blvd.", "Kirkland", "WA", "USA", "roland.mendel@xyz.com", "Janet.png"));
            employees.Add(new Employee("Margaret", "Peacock", "Sales Representative", "4110 Old Redmond Rd.", "Redmond", "WA", "USA", "margaret.peacock@xyz.com", "Margaret.png"));
            employees.Add(new Employee("Steven", "Buchanan", "Sales Manager", "14 Garrett Hill", "London", string.Empty, "UK", "steven.buchanan@xyz.com", "Steven.png"));
            return employees;
        }
        /// <summary>
        /// Represents a class to maintain employee details.
        /// </summary>
        public class Employee
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Address { get; set; }
            public string City { get; set; }
            public string Region { get; set; }
            public string Country { get; set; }
            public string Title { get; set; }
            public string Contact { get; set; }
            public string Photo { get; set; }
            public Employee(string firstName, string lastName, string title, string address, string city, string region, string country, string contact, string photoFilePath)
            {
                FirstName = firstName;
                LastName = lastName;
                Title = title;
                Address = address;
                City = city;
                Region = region;
                Country = country;
                Contact = contact;
                Photo = photoFilePath;
            }
        }
    }
}
