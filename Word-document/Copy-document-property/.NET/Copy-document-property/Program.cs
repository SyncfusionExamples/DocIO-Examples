using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Load an existing main Word document.
using (WordDocument targetDocument = new WordDocument(@"../../../Data/TargetDocument.docx", FormatType.Docx))
{
    //Load an existing template Word document.
    using (WordDocument sourceDocument = new WordDocument(@"../../../Data/SourceDocument.docx", FormatType.Docx))
    {
        //Move Built-in document properties from one Word document to another Word document.
        MoveBuiltinDocumentProperties(sourceDocument, targetDocument);
        //Move custom document properties from one Word document to another Word document.
        MoveCustomDocumentProperties(sourceDocument, targetDocument);
        //Save the Word document.
        targetDocument.Save(@"../../../Output/Result.docx");
    }
}

/// <summary>
/// Move Built-in document properties from one Word document to another Word document.
/// </summary>
void MoveBuiltinDocumentProperties(WordDocument sourceDocument, WordDocument targetDocument)
{
    if (sourceDocument.BuiltinDocumentProperties.Author != null)
        targetDocument.BuiltinDocumentProperties.Author = sourceDocument.BuiltinDocumentProperties.Author;

    if (sourceDocument.BuiltinDocumentProperties.Title != null)
        targetDocument.BuiltinDocumentProperties.Title = sourceDocument.BuiltinDocumentProperties.Title;

    if (sourceDocument.BuiltinDocumentProperties.Subject != null)
        targetDocument.BuiltinDocumentProperties.Subject = sourceDocument.BuiltinDocumentProperties.Subject;

    if (sourceDocument.BuiltinDocumentProperties.Keywords != null)
        targetDocument.BuiltinDocumentProperties.Keywords = sourceDocument.BuiltinDocumentProperties.Keywords;

    if (sourceDocument.BuiltinDocumentProperties.Comments != null)
        targetDocument.BuiltinDocumentProperties.Comments = sourceDocument.BuiltinDocumentProperties.Comments;

    if (sourceDocument.BuiltinDocumentProperties.Template != null)
        targetDocument.BuiltinDocumentProperties.Template = sourceDocument.BuiltinDocumentProperties.Template;

    if (sourceDocument.BuiltinDocumentProperties.LastAuthor != null)
        targetDocument.BuiltinDocumentProperties.LastAuthor = sourceDocument.BuiltinDocumentProperties.LastAuthor;

    if (sourceDocument.BuiltinDocumentProperties.Thumbnail != null)
        targetDocument.BuiltinDocumentProperties.Thumbnail = sourceDocument.BuiltinDocumentProperties.Thumbnail;

    if (sourceDocument.BuiltinDocumentProperties.ApplicationName != null)
        targetDocument.BuiltinDocumentProperties.ApplicationName = sourceDocument.BuiltinDocumentProperties.ApplicationName;

    if (sourceDocument.BuiltinDocumentProperties.Category != null)
        targetDocument.BuiltinDocumentProperties.Category = sourceDocument.BuiltinDocumentProperties.Category;

    if (sourceDocument.BuiltinDocumentProperties.Company != null)
        targetDocument.BuiltinDocumentProperties.Company = sourceDocument.BuiltinDocumentProperties.Company;

    if (sourceDocument.BuiltinDocumentProperties.Manager != null)
        targetDocument.BuiltinDocumentProperties.Manager = sourceDocument.BuiltinDocumentProperties.Manager;

    if (sourceDocument.BuiltinDocumentProperties.RevisionNumber != null)
        targetDocument.BuiltinDocumentProperties.RevisionNumber = sourceDocument.BuiltinDocumentProperties.RevisionNumber;

    if (sourceDocument.BuiltinDocumentProperties.CreateDate != null)
        targetDocument.BuiltinDocumentProperties.CreateDate = sourceDocument.BuiltinDocumentProperties.CreateDate;

    if (sourceDocument.BuiltinDocumentProperties.DocSecurity != -2147483648)
        targetDocument.BuiltinDocumentProperties.DocSecurity = sourceDocument.BuiltinDocumentProperties.DocSecurity;

    if (sourceDocument.BuiltinDocumentProperties.LastPrinted != null)
        targetDocument.BuiltinDocumentProperties.LastPrinted = sourceDocument.BuiltinDocumentProperties.LastPrinted;

    if (sourceDocument.BuiltinDocumentProperties.LastSaveDate != null)
        targetDocument.BuiltinDocumentProperties.LastSaveDate = sourceDocument.BuiltinDocumentProperties.LastSaveDate;

    if (sourceDocument.BuiltinDocumentProperties.TotalEditingTime != null)
        targetDocument.BuiltinDocumentProperties.TotalEditingTime = sourceDocument.BuiltinDocumentProperties.TotalEditingTime;
}

/// <summary>
/// Move custom document properties from one Word document to another Word document.
/// </summary>
void MoveCustomDocumentProperties(WordDocument sourceDocument, WordDocument targetDocument)
{
    for (int i = 0; i < sourceDocument.CustomDocumentProperties.Count; i++)
    {
        targetDocument.CustomDocumentProperties.Add(sourceDocument.CustomDocumentProperties[i].Name, sourceDocument.CustomDocumentProperties[i].Value);
    }
}