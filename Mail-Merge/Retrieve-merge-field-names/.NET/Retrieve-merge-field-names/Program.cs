using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Retrieve_merge_field_names
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Creates new Word document instance for Word processing.
                using (WordDocument document = new WordDocument())
                {
                    //Opens the Word template document.       
                    document.Open(fileStream, FormatType.Docx);

                    //Gets the merge field names from the document.
                    string[] fieldNames = document.MailMerge.GetMergeFieldNames();

                    Console.WriteLine("All merge field names in the Word document : ");
                    foreach (string fieldName in fieldNames)
                        Console.WriteLine(fieldName);

                    //Gets the merge field group names from the document.
                    string[] groupNames = document.MailMerge.GetMergeGroupNames();

                    Console.WriteLine("\n\nMerge field group names in the Word document : ");
                    foreach (string groupName in groupNames)
                        Console.WriteLine(groupName);

                    //Gets the fields from the specified groups. 
                    string[] fieldNamesInGroup = document.MailMerge.GetMergeFieldNames(groupNames[1]);

                    Console.WriteLine("\n\nMerge field names in " + groupNames[1] + " Group :");
                    foreach (string fieldNameInGroup in fieldNamesInGroup)
                        Console.WriteLine(fieldNameInGroup);
                }
            }
        }
    }
}
