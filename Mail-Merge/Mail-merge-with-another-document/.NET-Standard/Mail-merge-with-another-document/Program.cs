using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Mail_merge_with_another_document
{
    class Program
    {
        static int count = 0;
        static Dictionary<string, string> BookmarksAdded = new Dictionary<string, string>();
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Gets the subscription details as “IEnumerable” collection.
                    List<CategoryList> categoryLists = GetCategoryList();
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Categories", categoryLists);

                    //Mail merge event
                    document.MailMerge.MergeField += MailMerge_MergeField;

                    //Performs Mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataTable);

                    //Replace the field with another document content.
                    ReplaceBookmarks(document);

                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Data/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        private static void MailMerge_MergeField(object sender, MergeFieldEventArgs args)
        {
            if(args.FieldName == "Description")
            {
                //Get the owner paragraph
                WParagraph para = args.CurrentMergeField.OwnerParagraph;
                //Get the index of current field
                int index = para.Items.IndexOf(args.CurrentMergeField);
                //Add the bookmark to the begining of the field
                string bkmkName = "Bkmk_" + count;
                BookmarkStart bkmkStart = new BookmarkStart(para.Document, bkmkName);
                Bookmark bookmark = new Bookmark(bkmkStart);
                para.Items.Insert(index, bkmkStart);
                BookmarkEnd bkmkEnd = new BookmarkEnd(para.Document, bkmkName);
                para.Items.Insert(index + 1, bkmkEnd);
                //Add the bookmark name and the document path to the dictionary.
                BookmarksAdded.Add(bkmkName, args.FieldValue.ToString());
                //Increment the count
                count++;
                //Remove the field value i.e., document path given.
                args.Text = "";
               
            }
        }
        private static void ReplaceBookmarks(WordDocument document)
        {
            foreach(string bkmkName in BookmarksAdded.Keys)
            {
                using (FileStream fileStream = new FileStream(Path.GetFullPath(BookmarksAdded.GetValueOrDefault(bkmkName)), FileMode.Open, FileAccess.ReadWrite))
                {
                    //Opens the document for content.
                    using (WordDocument tempDoc = new WordDocument(fileStream, FormatType.Automatic))
                    {
                        //Get the bookmark using bookmark navigator
                        BookmarksNavigator nav = new BookmarksNavigator(document);
                        nav.MoveToBookmark(bkmkName);
                        //Load the document as WordDocumentPart
                        WordDocumentPart part = new WordDocumentPart();
                        part.Load(tempDoc);
                        //Replace the bookmark content.
                        nav.ReplaceContent(part);
                        //Remove the bookmark
                        Bookmark bkmk = document.Bookmarks.FindByName(bkmkName);
                        document.Bookmarks.Remove(bkmk);
                    }
                }
               
            }
        }

        #region Helper Methods
        public static List<CategoryList> GetCategoryList()
        {
            List<FieldList> field = new List<FieldList>();
            field.Add(new FieldList("Name", "Adventure Work Cycles"));
            field.Add(new FieldList("Product", "Bicycles"));
            field.Add(new FieldList("Location", "Washington"));

            List<ItemsList> items = new List<ItemsList>();
            items.Add(new ItemsList("Introduction", "../../../Data/One.html", field));

            field = new List<FieldList>();
            field.Add(new FieldList("Manufacturing plant", "Importadores Neptuno"));
            field.Add(new FieldList("Location", "Mexico"));
            field.Add(new FieldList("Year", "2000"));

            items.Add(new ItemsList("History", "../../../Data/Two.docx", field));

            List<CategoryList> categories = new List<CategoryList>();
            categories.Add(new CategoryList("Adventure Work Cycles", items));

            return categories;
        }
        #endregion
    }

    #region Helper class
    public class CategoryList
    {
        public string CategoryName { get; set; }
        public List<ItemsList> Items { get; set; }
        public CategoryList(string categoryName, List<ItemsList> items)
        {
            CategoryName = categoryName;
            Items = items;
        }
    }
    public class ItemsList
    {
        public string ItemTitle { get; set; }
        public string Description { get; set; }
        public List<FieldList> Field { get; set; }
        public ItemsList(string itemTitle, string description, List<FieldList> field)
        {
            ItemTitle = itemTitle;
            Description = description;
            Field = field;
        }
    }
    public class FieldList
    {
        public string FieldName { get; set; }
        public string FieldValue { get; set; }
        public FieldList(string name, string value)
        {
            FieldName = name;
            FieldValue = value;
        }
    }
    #endregion
}
