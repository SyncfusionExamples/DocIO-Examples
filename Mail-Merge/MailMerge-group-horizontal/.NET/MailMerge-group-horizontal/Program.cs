using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Collections.Generic;
using System.IO;

namespace MailMerge_group_horizontal
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the Person details as IEnumerable collection.
                    List<Person> PersonList = GetPersons();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("Person", PersonList);
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the Person details to perform mail merge.
        /// </summary>
        public static List<Person> GetPersons()
        {
            List<Person> Persons = new List<Person>();
            Persons.Add(new Person("Nancy"));
            Persons.Add(new Person("Andrew"));
            Persons.Add(new Person("Janet"));
            return Persons;
        }
    }

    /// <summary>
    /// Represents a class to maintain Person details.
    /// </summary>
    public class Person
    {
        public string FirstName { get; set; }

        public Person(string firstName)
        {
            FirstName = firstName;

        }
    }
}