using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Insert_merge_field_at_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load an existing Word document into DocIO instance
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    #region Insert paragraph at bookmark
                    // Create text bodypart
                    TextBodyPart bodyPart = new TextBodyPart(document);
                    // Create new paragraph and append merge field
                    WParagraph para = new WParagraph(document);
                    para.AppendField("Product", FieldType.FieldMergeField);
                    bodyPart.BodyItems.Add(para);

                    //Create the bookmark navigator instance to access the bookmark
                    BookmarksNavigator bkmk = new BookmarksNavigator(document);
                    //Move the virtual cursor to the location before the end of the bookmark
                    bkmk.MoveToBookmark("bookmark");
                    // Replace the bookmark content with our body part
                    bkmk.ReplaceBookmarkContent(bodyPart);
                    #endregion

                    #region Execute mailmerge
                    string[] fieldNames = { "Product", "ProductNo", "Size" };
                    string[] fieldValues = { "Cycle", "1234", "32" };
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    #endregion

                    //Creates file stream
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the loaded document in WordML format to the output stream
                        document.Save(outputFileStream, FormatType.Docx);
                        //Close the Word document
                        document.Close();
                    }
                }
            }
        }
    }
}
