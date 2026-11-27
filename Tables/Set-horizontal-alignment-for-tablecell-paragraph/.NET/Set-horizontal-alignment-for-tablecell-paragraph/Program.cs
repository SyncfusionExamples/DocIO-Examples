using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

using (FileStream inputStream = new FileStream("../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Open the template Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Retrieve the body of the first section.
        WTextBody body = document.Sections[0].Body;
        // Retrieve the first table from the body.
        WTable table = document.Sections[0].Tables[0] as WTable;
        foreach(WTableRow rows in table.Rows)
        {
            foreach(WTableCell cell in rows.Cells)
            {
                foreach(Entity ent in cell.ChildEntities)
                {
                    if(ent is WParagraph para)
                    {
                        para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                    }
                }
            }
        }
        // Save the modified document.
        using (FileStream outputStream = new FileStream("../../../Output/output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}