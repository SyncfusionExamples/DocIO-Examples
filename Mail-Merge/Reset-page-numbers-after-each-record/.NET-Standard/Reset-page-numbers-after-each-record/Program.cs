using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Reset_page_numbers_after_each_record
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the employee details as IEnumerable collection.
                    List<Employee> employeeList = GetEmployees();
                    //Creates an instance of MailMergeDataTable by specifying MailMerge group name and IEnumerable collection.
                    MailMergeDataTable dataSource = new MailMergeDataTable("Employees", employeeList); 
                    //Performs Mail merge.
                    document.MailMerge.ExecuteGroup(dataSource);

                    #region Split Word document by sections
                    //Find all the occurance of place holder - #SectionBreak#
                    TextSelection[] selections = document.FindAll("#SectionBreak#", false, false);
                    //Loop through all the text selection
                    for (int i = selections.Length - 1; i >= 0; i--)
                    {
                        TextSelection selection = selections[i];
                        //Get the owner paragraph of the selected text
                        WParagraph para = selection.GetAsOneRange().OwnerParagraph;
                        WSection srcSection = document.LastSection;
                        //Insert section break
                        InsertSectionBreak(para, srcSection);
                        WSection curSection = GetSection(para);
                        //Removes the place holder.
                        curSection.Body.ChildEntities.Remove(para);
                    }
                    #endregion

                    #region Resets the page number
                    //Iterates each section from Word document.
                    foreach (WSection section in document.Sections)
                    {
                        //Resets the page number.
                        section.PageSetup.RestartPageNumbering = true;
                        section.PageSetup.PageStartingNumber = 1;
                    }
                    #endregion
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        private static WSection InsertSectionBreak(TextBodyItem bodyItem, WSection srcSection)
        {
            //Get the current section of the body item.
            var currentSection = GetSection(bodyItem);

            // Identify all the items in the same Section that are positioned after the bodyItem. These body items need to be cut and pasted to the new section.
            int numBodyItemsToStay = GetIndex(bodyItem) + 1;
            var entityCollection = currentSection.Body.ChildEntities;
            var bodyItemsToMove = entityCollection.Cast<TextBodyItem>()
                                              .Skip(numBodyItemsToStay)
                                              .ToList();

            //Create a new section that is positioned after the current section.
            var newSection = new WSection(bodyItem.Document);
            //Add new section as a sibling of current section.
            AddSiblings(currentSection, new[] { newSection });

            // Cut and paste each marked body item from the current section to the new section.
            foreach (var bodyItemToMove in bodyItemsToMove)
            {
                newSection.Body.ChildEntities.Add(bodyItemToMove);
            }
            return newSection;
        }
        /// <summary>
        /// Get the index of the particular entity.
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        private static int GetIndex(IEntity entity)
        {
            ICompositeEntity container = entity.Owner as ICompositeEntity;
            if (container == null)
            {
                throw new ApplicationException("Entity is not index-able as it does not have a valid container.");
            }

            return container.ChildEntities.IndexOf(entity);
        }
        /// <summary>
        /// Geth the section of the specified entity.
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        private static WSection GetSection(IEntity  entity)
        {
            if (entity is WSection)
            {
                return (WSection)entity;
            }

            if (entity is WordDocument)
            {
                throw new ApplicationException("WordDocument does not belong to any sections.");
            }
            // Traverse the tree bottom-up until the Section is found.
            IEntity parentEntity = entity.Owner;
            while (parentEntity != null)
            {
                if (parentEntity is WSection)
                {
                    return (WSection)parentEntity;
                }

                parentEntity = parentEntity.Owner;
            }

            // Unable to find the Section this entity belongs to. This entity is most likely not attached to any containers yet.
            return null;
        }
        /// <summary>
        /// Add new section as sibling of current section.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="entity"></param>
        /// <param name="newSiblings"></param>
        public static void AddSiblings<T>(IEntity entity, IEnumerable<T> newSiblings) where T : class, IEntity
        {
            int newIndex = GetIndex(entity) + 1;

            ICompositeEntity container = entity.Owner as ICompositeEntity;
            if (container == null)
            {
                throw new ApplicationException("Unable to add new siblings to this entity as it does not have a valid container.");
            }

            foreach (var newSibling in newSiblings)
            {
                container.ChildEntities.Insert(newIndex++, newSibling);
            }
        }

        /// <summary>
        /// Gets the employee details to perform mail merge.
        /// </summary>
        public static List<Employee> GetEmployees()
        {
            List<Employee> employees = new List<Employee>();
            employees.Add(new Employee("Nancy", "Smith", "1","Sales Representative", "505 - 20th Ave. E. Apt. 2A,", "Seattle", "WA", "USA"));
            employees.Add(new Employee("Andrew", "Fuller","2", "Vice President, Sales", "908 W. Capital Way", "Tacoma", "WA", "USA"));
            employees.Add(new Employee("Roland", "Mendel","3", "Sales Representative", "722 Moss Bay Blvd.", "Kirkland", "WA", "USA"));
            return employees;
        }
    }
    /// <summary>
    /// Represents a class to maintain employee details.
    /// </summary>
    public class Employee
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmployeeID {get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string Country { get; set; }
        public string Title { get; set; }
        public string Photo { get; set; }
        public Employee(string firstName,  string lastName, string employeeId, string title, string address, string city, string region, string country)
        {
            FirstName = firstName;
            LastName = lastName;
            EmployeeID = employeeId;
            Title = title;
            Address = address;
            City = city;
            Region = region;
            Country = country;
        }
    }
}
