using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Find_and_replace_merge_field_name
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Opens the template Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    // Finds all merge fields in the document.
                    List<Entity> mergeFields = document.FindAllItemsByProperty(EntityType.MergeField, null, null);
                    // Replaces the merge field name "first_name" with "FirstName".
                    ReplaceMergeFieldName("first_name", "FirstName", mergeFields);
                    // Replaces the merge field name "last_name" with "LastName".
                    ReplaceMergeFieldName("last_name", "LastName", mergeFields);
                    // Creates a file stream to save the modified document.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        // Saves the Word document to the file stream in DOCX format.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
		
		/// <summary>
        /// Replaces the specified merge field name with a new name in the given list of merge fields.
        /// </summary>
        private static void ReplaceMergeFieldName(string nameToFind, string nameToReplace, List<Entity> mergeFields)
        {
            // Iterates through the list of merge fields.
            foreach (Entity field in mergeFields)
            {
                // Checks if the current merge field matches the name to be replaced.
                if ((field as WMergeField).FieldName == nameToFind)
                {
                    // Updates the merge field name to the new name.
                    (field as WMergeField).FieldName = nameToReplace;
                }
            }
        }
    }
}
