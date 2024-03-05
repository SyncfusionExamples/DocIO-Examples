using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;


//Load an existing Word document.
using FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read);
using WordDocument document = new WordDocument(fileStream, FormatType.Docx);
WParagraph paragraph = new WParagraph(document);
paragraph.AppendText("List of Figures");
//Apply Heading1 style for paragraph.
paragraph.ApplyStyle(BuiltinStyle.Heading1);
//Insert the paragraph 
document.LastSection.Body.ChildEntities.Insert(0, paragraph);
//Create new paragraph and append TOC 
paragraph = new WParagraph(document);
TableOfContent tableOfContent = paragraph.AppendTOC(1, 3);
//Disable a flag to exclude heading style paragraphs in TOC entries.
tableOfContent.UseHeadingStyles = false;
//Set the name of SEQ field identifier for table of figures.
tableOfContent.TableOfFiguresLabel = "Figure";
//Insert the paragraph to the text body.
document.LastSection.Body.ChildEntities.Insert(1, paragraph);

//Find all pictures from the document
List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);
// Iterate each picture and add caption.
foreach (WPicture picture in pictures)
{
    //Set alternate text as caption for picture.
    WParagraph captionPara = picture.AddCaption("Figure", CaptionNumberingFormat.Number, CaptionPosition.AfterImage) as WParagraph;
    captionPara.AppendText(" " + picture.AlternativeText);
    //Apply formatting to the caption.
    captionPara.ApplyStyle(BuiltinStyle.Caption);
    captionPara.ParagraphFormat.BeforeSpacing = 8;
    captionPara.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
}

// Create a new paragraph
paragraph = new WParagraph(document);
paragraph.AppendText("List of Tables");
// Apply Heading1 style for paragraph.
paragraph.ApplyStyle(BuiltinStyle.Heading1);
// Insert the paragraph
document.LastSection.Body.ChildEntities.Insert(2, paragraph);

//Create a new paragraph and append TOC.
paragraph = new WParagraph(document);
tableOfContent = paragraph.AppendTOC(1, 3);
//Disable a flag to exclude heading style paragraphs in TOC entries.
tableOfContent.UseHeadingStyles = false;
//Set the name of SEQ field identifier for table of tables.
tableOfContent.TableOfFiguresLabel = "Table";
// Insert the paragraph to the text body.
document.LastSection.Body.ChildEntities.Insert(3, paragraph);


// Find all tables from the document
List<Entity> tables = document.FindAllItemsByProperty(EntityType.Table, null, null);
//Iterate each table and add caption.
foreach (WTable table in tables)
{
    //Gets the table index
    int tableIndex = table.OwnerTextBody.ChildEntities.IndexOf(table);
    //Create a new paragraph and appends the sequence field to use as a caption.
    WParagraph captionPara = new WParagraph(document);
    captionPara.AppendText("Table ");
    captionPara.AppendField("Table", FieldType.FieldSequence);
    //Set alternate text as caption for table.
    captionPara.AppendText(" " + table.Description);
    // Apply formatting to the paragraph
    captionPara.ApplyStyle(BuiltinStyle.Caption);
    captionPara.ParagraphFormat.BeforeSpacing = 8;
    captionPara.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
    //Insert the paragraph next to the table
    table.OwnerTextBody.ChildEntities.Insert(tableIndex + 1, captionPara);
}

//Update all document fields to update SEQ fields.
document.UpdateDocumentFields();
//Update the table of contents.
document.UpdateTableOfContents();

//Create a FileStream to save the Word document.
using FileStream outputStream = new FileStream("Result.docx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
//Save the Word document.
document.Save(outputStream, FormatType.Docx);