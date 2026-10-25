using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;


WordDocument document = new WordDocument(@"../../../Data/Template", FormatType.Docx);

#region InsertTOC
//Create a new paragraph and append TOC.
WParagraph paragraph = new WParagraph(document);
//Create TOC that includes Heading 1 to Heading 3 style
TableOfContent tableOfContent = paragraph.AppendTOC(1, 3);
//Disable a flag to exclude heading style paragraphs in TOC entries.
tableOfContent.UseHeadingStyles = true;
//Enable to use TOC entry fields
tableOfContent.UseTableEntryFields = true;
//Insert paragraph with TOC at first of the document
document.Sections[0].Body.ChildEntities.Insert(0, paragraph);
#endregion

#region InsertTOCEntryField
//Access exisiting paragraph
paragraph = (WParagraph)(document.Sections[0].Tables[1].Rows[0].Cells[1].ChildEntities[0] as WTable).Rows[0].Cells[0].ChildEntities[0];
//Append field of type TOCEntry
WField field = (WField)paragraph.AppendField(paragraph.Text, FieldType.FieldTOCEntry);
#endregion

#region ApplyHeadingStyle
paragraph = (WParagraph)(document.Sections[0].Tables[1].Rows[0].Cells[1].ChildEntities[0] as WTable).Rows[1].Cells[0].ChildEntities[0];
//Apply "Heading 1" style for paragraph
paragraph.ApplyStyle(BuiltinStyle.Heading1);
//paragraph.ApplyStyle(BuiltinStyle.Heading2);
#endregion

//Update Table of contents
document.UpdateTableOfContents();

//Save updated document
document.Save(@"../../../Output/Result.docx", FormatType.Docx);