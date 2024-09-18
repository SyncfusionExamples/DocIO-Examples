// See https://aka.ms/new-console-template for more information
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (WordDocument document = new WordDocument())
{
    //Opens the Word template document.
    Stream docStream = File.OpenRead(Path.GetFullPath(@"Data/TOC.docx"));
    document.Open(docStream, FormatType.Docx);
    docStream.Dispose();
    //Removes the TOC field.
    TableOfContent toc = document.Sections[0].Body.Paragraphs[2].Items[0] as TableOfContent;
    RemoveTableOfContents(toc);
    //Saves the file in the given path
    docStream = File.Create(Path.GetFullPath(@"Output/Sample.docx"));
    document.Save(docStream, FormatType.Docx);
    docStream.Dispose();
}

#region Helper methods
/// <summary>
/// Removes the table of contents from Word document.
/// </summary>
/// <param name="toc"></param>
static void RemoveTableOfContents(TableOfContent toc)
{
    //Finds the last TOC item.
    Entity lastItem = FindLastTOCItem(toc);

    //TOC field end mark wasn't exist.
    if (lastItem == null)
        return;

    //Inserts the bookmark start before the TOC instance.
    BookmarkStart bkmkStart = new BookmarkStart(toc.Document, "tableOfContent");
    toc.OwnerParagraph.Items.Insert(toc.OwnerParagraph.Items.IndexOf(toc), bkmkStart);

    //Inserts the bookmark end to next of TOC last item.
    BookmarkEnd bkmkEnd = new BookmarkEnd(toc.Document, "tableOfContent");
    WParagraph paragraph = lastItem.Owner as WParagraph;
    paragraph.Items.Insert(paragraph.Items.IndexOf(lastItem) + 1, bkmkEnd);

    //Delete all the items from bookmark start to end (TOC items) using Bookmark Navigator.
    DeleteBookmarkContents(bkmkEnd.Name, toc.Document);
}
/// <summary>
/// Delete the bookmark items.
/// </summary>
/// <param name="bkmkName"></param>
/// <param name="document"></param>
static void DeleteBookmarkContents(string bkmkName, WordDocument document)
{
    //Creates the bookmark navigator instance to access the bookmark
    BookmarksNavigator navigator = new BookmarksNavigator(document);
    //Moves the virtual cursor to the location before the end of the bookmark "tableOfContent".
    navigator.MoveToBookmark(bkmkName);
    //Deletes the bookmark content.
    navigator.DeleteBookmarkContent(false);
    //Gets the bookmark instance by using FindByName method of BookmarkCollection with bookmark name.
    Bookmark bookmark = document.Bookmarks.FindByName(bkmkName);
    //Removes the bookmark named "tableOfContent" from Word document.
    document.Bookmarks.Remove(bookmark);
}
/// <summary>
/// Finds the last TOC item.
/// </summary>
/// <param name="toc"></param>
/// <returns></returns>
static Entity FindLastTOCItem(TableOfContent toc)
{
    int tocIndex = toc.OwnerParagraph.Items.IndexOf(toc);
    //TOC may contains nested fields and each fields has its owner field end mark 
    //so to identify the TOC Field end mark (WFieldMark instance) used the stack.
    Stack<Entity> fieldStack = new Stack<Entity>();
    fieldStack.Push(toc);

    //Finds whether TOC end item is exist in same paragraph.
    for (int i = tocIndex + 1; i < toc.OwnerParagraph.Items.Count; i++)
    {
        Entity item = toc.OwnerParagraph.Items[i];

        if (item is WField)
            fieldStack.Push(item);
        else if (item is WFieldMark && (item as WFieldMark).Type == FieldMarkType.FieldEnd)
        {
            if (fieldStack.Count == 1)
            {
                fieldStack.Clear();
                return item;
            }
            else
                fieldStack.Pop();
        }
    }
    return FindLastItemInTextBody(toc, fieldStack);
}
/// <summary>
/// Finds the last TOC item from consequence text body items.
/// </summary>
/// <param name="toc"></param>
/// <param name="fieldStack"></param>
/// <returns></returns>
static Entity FindLastItemInTextBody(TableOfContent toc, Stack<Entity> fieldStack)
{
    WTextBody tBody = toc.OwnerParagraph.OwnerTextBody;

    //Finds whether TOC end item is exist in text body items.
    for (int i = tBody.ChildEntities.IndexOf(toc.OwnerParagraph) + 1; i < tBody.ChildEntities.Count; i++)
    {
        WParagraph paragraph = null;
        if (tBody.ChildEntities[i] is WParagraph)
            paragraph = tBody.ChildEntities[i] as WParagraph;
        else
            continue;

        foreach (Entity item in paragraph.Items)
        {
            if (item is WField)
                fieldStack.Push(item);
            else if (item is WFieldMark && (item as WFieldMark).Type == FieldMarkType.FieldEnd)
            {
                if (fieldStack.Count == 1)
                {
                    fieldStack.Clear();
                    return item;
                }
                else
                    fieldStack.Pop();
            }
        }
    }
    return null;
}
#endregion
