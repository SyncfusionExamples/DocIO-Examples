using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_list_restart_numbering_HTML_with_text
{
    internal class Program
    {
        // List to store the names of different list styles used in the document.
        static List<string> listStyleNames = new List<string>();
        static void Main(string[] args)
        {
            // Load the input Word document from file stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Open the Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Load the input Word document from file stream
                    using (FileStream htmlStream = new FileStream(Path.GetFullPath(@"Data/sample.html"), FileMode.Open, FileAccess.Read))
                    {
                        // Open the Word document
                        using (WordDocument replaceDoc = new WordDocument(htmlStream, FormatType.Html))
                        {
                            //Replace the first word with HTML file content
                            ReplaceText(document, "Tag1", replaceDoc, true);
                            //Replace the second word with HTML file content
                            ReplaceText(document, "Tag2", replaceDoc, false);

                            // Save the modified document to a new file
                            using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                            {
                                document.Save(docStream1, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
        private static void ReplaceText(WordDocument document, string findText, WordDocument replaceDoc, bool isFirstReplace)
        {
            TextSelection selection = document.Find(findText, true, true);
            if (selection != null)
            {
                //Get the textrange
                WTextRange textRange = selection.GetAsOneRange();
                //Get the owner paragraph
                WParagraph ownerPara = textRange.OwnerParagraph;

                //For first time replacement alone.
                if (isFirstReplace)
                {
                    //Get the index of the textrange
                    int index = ownerPara.ChildEntities.IndexOf(textRange);
                    //Add the bookmark start before the textrange
                    BookmarkStart bookmarkStart = new BookmarkStart(document, "Bkmk");
                    ownerPara.ChildEntities.Insert(index, bookmarkStart);
                    //Increment the index
                    index++;
                    //Add the bookmark end after the textrange
                    BookmarkEnd bookmarkEnd = new BookmarkEnd(document, "Bkmk");
                    ownerPara.ChildEntities.Insert(index + 1, bookmarkEnd);
                    //Replace the text with HTML content
                    document.Replace(findText, replaceDoc, true, true);
                    //Navigate to the bookmark content
                    BookmarksNavigator navigator = new BookmarksNavigator(document);
                    navigator.MoveToBookmark("Bkmk");
                    //Get the bookmark content
                    TextBodyPart bodyPart = navigator.GetBookmarkContent();
                    //Get the list of list styles
                    GetListStyleName(bodyPart.BodyItems);
                    //Remove the bookmark
                    Bookmark bookmark = document.Bookmarks.FindByName("Bkmk");
                    document.Bookmarks.Remove(bookmark);
                }
                else
                {
                    //Get the next sibiling of the owner paragraph
                    IEntity nextSibiling = ownerPara.NextSibling;
                    //Get the owner paragraph index as start index 
                    int startIndex = ownerPara.OwnerTextBody.ChildEntities.IndexOf(ownerPara);
                    //Replace the text with HTML content
                    document.Replace(findText, replaceDoc, true, true);
                    //Get the end index
                    //If the next sibiling is present then it is the end index, else the child entities count
                    int endIndex = nextSibiling != null ? ownerPara.OwnerTextBody.ChildEntities.IndexOf(nextSibiling)
                        : ownerPara.OwnerTextBody.ChildEntities.Count;
                    //Restart the numbering
                    RestartNumbering(startIndex, endIndex, document.Sections[0].Body.ChildEntities);
                }
            }
        }
        //Get the list style names from the collection
        private static void GetListStyleName(EntityCollection collection)
        {
            //Iterate through the collection
            foreach (Entity entity in collection)
            {
                switch (entity.EntityType)
                {
                    //Entity is paragraph
                    case EntityType.Paragraph:
                        WParagraph wParagraph = (WParagraph)entity;
                        //Check whether the paragrah has list format with numbered type which is not in the collection list.
                        if (wParagraph.ListFormat.CurrentListLevel != null
                            && wParagraph.ListFormat.ListType == ListType.Numbered
                            && !listStyleNames.Contains(wParagraph.ListFormat.CurrentListStyle.Name))
                            //Add the list style name to the collection list
                            listStyleNames.Add(wParagraph.ListFormat.CurrentListStyle.Name);
                        break;
                    //Entity is Table
                    case EntityType.Table:
                        WTable table = (WTable)entity;
                        //Iterate thorugh rows
                        foreach (WTableRow row in table.Rows)
                        {
                            //Iterate through cells
                            foreach (WTableCell cell in row.Cells)
                            {
                                //Get the list style name
                                GetListStyleName(cell.ChildEntities);
                            }
                        }
                        break;
                }
            }
        }
        //Restart the numbering for replaced items
        private static void RestartNumbering(int startIndex, int endIndex, EntityCollection collection)
        {
            //Local value
            string listName = string.Empty;
            for (int i = startIndex; i < endIndex; i++)
            {
                Entity entity = collection[i];
                switch (entity.EntityType)
                {
                    //Entity is Paragraph
                    case EntityType.Paragraph:
                        WParagraph wParagraph = (WParagraph)entity;
                        //Check whether the paragraph have current list level and the same list name in the collection list.
                        if (wParagraph.ListFormat.CurrentListLevel != null
                            && listStyleNames.Contains(wParagraph.ListFormat.CurrentListStyle.Name))
                        {
                            //If the local name is not equal to current list name, then restart the numbering
                            if (listName != wParagraph.ListFormat.CurrentListStyle.Name)
                            {
                                //Set the current list name as local name
                                listName = wParagraph.ListFormat.CurrentListStyle.Name;
                                //Enable restart numbering
                                wParagraph.ListFormat.RestartNumbering = true;
                            }
                            //If the local name and current list name are equal then continue list numbering
                            else
                                wParagraph.ListFormat.ContinueListNumbering();
                        }
                        break;
                    //Entity is table
                    case EntityType.Table:
                        WTable table = (WTable)entity;
                        //Iterate through rows
                        foreach (WTableRow row in table.Rows)
                        {
                            //Iterate thorugh cells
                            foreach (WTableCell cell in row.Cells)
                            {
                                //Restart numbering for child entities in the cell.
                                RestartNumbering(0, cell.ChildEntities.Count, cell.ChildEntities);
                            }
                        }
                        break;
                }
            }
        }
    }
}
