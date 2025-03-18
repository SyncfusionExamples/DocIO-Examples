using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Modify_font_during_mail_merge
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NMaF5cXmBCf1FpRmJGdld5fUVHYVZUTXxaS00DNHVRdkdmWX1cdnRRQ2NcUkZwXUo=");
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    // Retrieve all merge field group names present in the document.
                    string[] mergeGroupNames = document.MailMerge.GetMergeGroupNames();
                    //Modify group names in the document
                    ModifyGroupName(document, mergeGroupNames);
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper methods
        /// <summary>
        /// Find and modify merge group name in the document
        /// </summary>
        /// <param name="document"></param>
        /// <param name="mergeGroupNames"></param>
        private static void ModifyGroupName(WordDocument document, string[] mergeGroupNames)
        {
            // Iterate through each merge group name.
            foreach (string mergeGroupName in mergeGroupNames)
            {
                // Find the merge field in the document that matches the current group name (Contains begin and end group).
                // The "FieldName" property is used to locate the specific merge field.
                List<Entity> mergeFieldsWithFieldName = document.FindAllItemsByProperty(EntityType.MergeField, "FieldName", mergeGroupName);
                foreach (Entity entity in mergeFieldsWithFieldName)
                {
                    WMergeField currentMergeField = entity as WMergeField;
                    // You can modify the field name as needed.
                    // Example: Modify the field name by appending "_123" to it.
                    currentMergeField.FieldName = mergeGroupName + "_123";
                }
            }
        }
        #endregion
    }
}
