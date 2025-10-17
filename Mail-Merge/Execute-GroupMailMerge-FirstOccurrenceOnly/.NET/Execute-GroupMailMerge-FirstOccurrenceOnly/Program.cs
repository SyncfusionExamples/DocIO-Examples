using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Execute_GroupMailMerge_FirstOccurrenceOnly
{
    class Program
    {
        static void Main()
        {
            // Load the Word template document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx"));
            // Get all merge field names from the document
            string[] mergeGroupNames = document.MailMerge.GetMergeFieldNames();
            // Create a dictionary to count how many times each merge field appears
            Dictionary<string, int> groupNameCounts = new Dictionary<string, int>();
            foreach (string groupName in mergeGroupNames)
            {
                // Increase count if field already exists, else add it
                if (groupNameCounts.ContainsKey(groupName))
                    groupNameCounts[groupName]++;
                else
                    groupNameCounts[groupName] = 1;
            }
            // Remove duplicate merge field groups
            foreach (var groupEntry in groupNameCounts)
            {
                string groupName = groupEntry.Key;
                // Skip if the field appears only once
                if (groupEntry.Value <= 1)
                    continue;
                // Find all merge fields with the same name
                List<Entity> mergeGroups = document.FindAllItemsByProperty(EntityType.MergeField, "FieldName", groupName);
                // Start from second occurrence to remove duplicates
                for (int i = 1; i < mergeGroups.Count; i++)
                {
                    WMergeField mergeField = mergeGroups[i] as WMergeField;
                    // Check if it's a group start field
                    if (mergeField.FieldCode.Contains("TableStart") || mergeField.FieldCode.Contains("BeginGroup"))
                    {
                        // Add bookmark start before the group
                        BookmarkStart bkmkStart = new BookmarkStart(document, groupName);
                        WParagraph startPara = mergeField.OwnerParagraph;
                        int mergeFieldIndex = startPara.ChildEntities.IndexOf(mergeField);
                        startPara.ChildEntities.Insert(mergeFieldIndex, bkmkStart);
                        // Add bookmark end after the group
                        WMergeField endField = mergeGroups[i + 1] as WMergeField;
                        BookmarkEnd bkmkEnd = new BookmarkEnd(document, groupName);
                        WParagraph endPara = endField.OwnerParagraph;
                        int endFieldIndex = endPara.ChildEntities.IndexOf(endField);
                        endPara.ChildEntities.Insert(endFieldIndex + 1, bkmkEnd);
                        // Delete content inside the bookmark
                        BookmarksNavigator navigator = new BookmarksNavigator(document);
                        navigator.MoveToBookmark(groupName);
                        navigator.DeleteBookmarkContent(false);
                        document.Bookmarks.Remove(navigator.CurrentBookmark);
                        // Remove owner table if applicable
                        if (startPara.OwnerTextBody.Owner.EntityType == EntityType.TableRow)
                        {
                            WTableRow tableRow = startPara.OwnerTextBody.Owner as WTableRow;
                            WTable ownerTable = tableRow.Owner as WTable;
                            Entity currentEntity = ownerTable;
                            // Traverse up to find the section and remove the table
                            while (currentEntity != null)
                            {
                                if (currentEntity is WSection section)
                                {
                                    section.Body.ChildEntities.Remove(ownerTable);
                                    break;
                                }
                                currentEntity = currentEntity.Owner;
                            }
                        }
                    }
                }
            }
            // Prepare nested mail merge data
            List<Organization> organizationList = GetOrganizations();
            MailMergeDataTable dataTable = new MailMergeDataTable("Organizations", organizationList);
            // Execute nested mail merge using the data
            document.MailMerge.ExecuteNestedGroup(dataTable);
            // Save the result
            document.Save(Path.GetFullPath(@"../../../Output/Result.docx"));
            // Close the document
            document.Close();
        }
    
        /// <summary>
        /// Create sample organization and employee data
        /// </summary>
        /// <returns>Return the organization list</returns>
        public static List<Organization> GetOrganizations()
        {
            // Create a list of employees
            List<EmployeeDetails> employees = new List<EmployeeDetails>
            {
                new EmployeeDetails("Thomas Hardy", "1001", "05/27/1996"),
                new EmployeeDetails("Maria Anders", "1002", "04/10/1998")
            };
            // Create a list of organizations with employee data
            List<Organization> organizations = new List<Organization>
            {
                new Organization("UK Office", "120 Hanover Sq.", "UK", employees)
            };
            // Return the organization list
            return organizations;
        }
    }
}
