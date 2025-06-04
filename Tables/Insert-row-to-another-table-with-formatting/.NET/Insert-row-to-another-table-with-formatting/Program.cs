using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Load the Word document from the specified path.
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Access the first table in the last section of the document (source table).
    WTable sourceTable = (WTable)document.LastSection.Tables[0];

    // Access the second table in the last section of the document (target table).
    WTable targetTable = (WTable)document.LastSection.Tables[1];

    // Clone the first row from the source table and add it to the target table as first row.
    targetTable.Rows.Insert(0, sourceTable.Rows[0].Clone());

    // Save the modified document to the specified output path in DOCX format.
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}
