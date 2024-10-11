using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

// Creates a new instance of WordDocument (empty Word document).
WordDocument wordDocument = new WordDocument();
// Adds a section to the document.
IWSection section = wordDocument.AddSection();
// Adds a paragraph to the section.
IWParagraph paragraph = section.AddParagraph();
// Adds a document variable named "Barcode" with a value of "123456".
wordDocument.Variables.Add("Barcode", "123456");
// Appends a field to the paragraph. The field type is set as FieldUnknown initially.
WField field = paragraph.AppendField("DISPLAYBARCODE ", Syncfusion.DocIO.FieldType.FieldUnknown) as WField;
// Specifies the field code for the barcode.
InsertBarcodeField(paragraph, field);
// Updates all document fields to reflect changes.
wordDocument.UpdateDocumentFields();
// Saves the output Word document to a specified file stream in DOCX format.
FileStream stream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate);
wordDocument.Save(stream, FormatType.Docx);
// Closes the file stream to release the resource.
stream.Close();

/// <summary>
/// Inserts the field code with a nested field.
/// </summary>
static void InsertBarcodeField(IWParagraph paragraph, WField? field)
{
    // Get the index of the field in the paragraph to insert the field code.
    int fieldIndex = paragraph.Items.IndexOf(field) + 1;

    // Add the field code for the barcode.
    field.FieldCode = "DISPLAYBARCODE ";

    // Increment field index for further insertion.
    fieldIndex++;

    // Insert document variables into the field code at the specified index.
    InsertDocVariables("Barcode", ref fieldIndex, paragraph);

    // Insert the barcode type (e.g., CODE128) into the paragraph.
    InsertText(" CODE128", ref fieldIndex, paragraph as WParagraph);
}

/// <summary>
/// Inserts text such as operators or identifiers into the paragraph.
/// </summary>
static void InsertText(string text, ref int fieldIndex, WParagraph paragraph)
{
    // Create a new text range with the specified text.
    WTextRange textRange = new WTextRange(paragraph.Document);
    textRange.Text = text;

    // Insert the text range as a field code item at the specified index.
    paragraph.Items.Insert(fieldIndex, textRange);

    // Increment the field index for subsequent insertions.
    fieldIndex++;
}

/// <summary>
/// Inserts document variables at the given index in the specified paragraph.
/// </summary>
static void InsertDocVariables(string fieldName, ref int fieldIndex, IWParagraph paragraph)
{
    // Create a new paragraph to hold the document variable field.
    WParagraph para = new WParagraph(paragraph.Document);
    para.AppendField(fieldName, FieldType.FieldDocVariable);

    // Get the count of child entities in the paragraph (should be 1 for the field).
    int count = para.ChildEntities.Count;

    // Insert the field into the original paragraph, ensuring the complete field structure is added.
    paragraph.ChildEntities.Insert(fieldIndex, para.ChildEntities[0]);

    // Increment the field index based on the count of inserted entities.
    fieldIndex += count;
}
