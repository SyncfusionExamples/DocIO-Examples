using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replacing_MergeField_With_Bullet_List
{
    class Program
    {
        static void Main(string[] args)
        {
            // Loads existing Word document
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
            {
                string[] fieldNames = { "ContactName", "CompanyName" };
                string[] fieldValues = { "Nancy", "Syncfusion" };

                // Uses mail merge events to perform formatting during runtime
                wordDocument.MailMerge.MergeField += new MergeFieldEventHandler(ApplyListFormat);

                // Performs mail merge
                wordDocument.MailMerge.Execute(fieldNames, fieldValues);

                // Saves Word document
                wordDocument.Save(Path.GetFullPath(@"Output/Output.docx"));
            }
        }
        
        /// <summary>
        /// Applies the default bullet list formatting to the paragraph
        /// that contains the current mail merge field.
        /// </summary>
        /// <param name="sender">The source of the mail merge event.</param>
        /// <param name="args">
        /// Contains data related to the current merge field, including
        /// the merge field and its owner paragraph.
        /// </param>
        private static void ApplyListFormat(object sender, MergeFieldEventArgs args)
        {
            // Gets the owner paragraph of the current merge field
            WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;

            // Applies bullet list style to the paragraph
            paragraph.ListFormat.ApplyDefBulletStyle();
        }

    }
}