using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Reset_numbering_for_SEQ_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a Word document.
            using (WordDocument document = CreateDocument())
            {
                //Accesses sequence field in the document.
                IWTable table = document.LastSection.Body.ChildEntities[1] as WTable;
                WSeqField field = ((table[0, 1].ChildEntities[0] as WParagraph).ChildEntities[1] as WSeqField);
                //Resets the number for sequence field.
                field.ResetNumber = 1001;
                //Accesses sequence field in the document.
                field = ((table[1, 1].ChildEntities[0] as WParagraph).ChildEntities[1] as WSeqField);
                //Resets the number for sequence field.
                field.ResetNumber = 1002;
                //Accesses sequence field in the document.
                field = ((table[2, 1].ChildEntities[0] as WParagraph).ChildEntities[1] as WSeqField);
                //Resets the number for sequence field.
                field.ResetNumber = 1003;
                //Accesses sequence field in the document.
                table = document.LastSection.Body.ChildEntities[3] as WTable;
                field = ((table[0, 1].ChildEntities[1] as WParagraph).ChildEntities[1] as WSeqField);
                //Resets the heading level for sequence field.
                field.ResetHeadingLevel = 1;
                //Updates the document fields.
                document.UpdateDocumentFields();
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        /// <summary>
        /// Creates a new Word document.
        /// </summary>
        private static WordDocument CreateDocument()
        {
			//Creates a new word document.
			WordDocument document = new WordDocument();
			//Adds new section to the document.
			IWSection section = document.AddSection();
			//Sets margin of the section.
			section.PageSetup.Margins.All = 72;
			//Adds new paragraph to the section.
			IWParagraph paragraph = section.AddParagraph() as WParagraph;
			//Adds text range.
			IWTextRange textRange = paragraph.AppendText("Adventure Works Cycles");
			textRange.CharacterFormat.FontSize = 16;
			textRange.CharacterFormat.Bold = true;
			paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
			//Adds a new table into Word document.
			IWTable table = section.AddTable();
			//Specifies the total number of rows & columns.
			table.ResetCells(3, 2);
			//First row.
			FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Nancy.png"), FileMode.Open, FileAccess.ReadWrite);
			table[0, 0].AddParagraph().AppendPicture(imageStream);
			paragraph = table[0, 1].AddParagraph();
			paragraph.AppendText("Employee Id: ");
			paragraph.AppendField("Id", FieldType.FieldSequence);
			table[0, 1].AddParagraph().AppendText("Name: Nancy Davolio");
			table[0, 1].AddParagraph().AppendText("Title: Sales Representative");
			table[0, 1].AddParagraph().AppendText("Address: 507 - 20th Ave. E.");
			table[0, 1].AddParagraph().AppendText("Zip Code: 98122");
			//Second row.
			imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Andrews.png"), FileMode.Open, FileAccess.ReadWrite);
			table[1, 0].AddParagraph().AppendPicture(imageStream);
			paragraph = table[1, 1].AddParagraph();
			paragraph.AppendText("Employee ID: ");
			paragraph.AppendField("Id", FieldType.FieldSequence);
			table[1, 1].AddParagraph().AppendText("Name: Andrew Fuller");
			table[1, 1].AddParagraph().AppendText("Title: Vice President, Sales");
			table[1, 1].AddParagraph().AppendText("Address1: 908 W. Capital Way, ");
			table[1, 1].AddParagraph().AppendText("TacomaWA USA");
			//Third row.
			imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Janet.png"), FileMode.Open, FileAccess.ReadWrite);
			table[2, 0].AddParagraph().AppendPicture(imageStream);
			paragraph = table[2, 1].AddParagraph();
			paragraph.AppendText("Employee ID: ");
			paragraph.AppendField("Id", FieldType.FieldSequence);
			table[2, 1].AddParagraph().AppendText("Name: Janet Leverling");
			table[2, 1].AddParagraph().AppendText("Title: Sales Representative");
			table[2, 1].AddParagraph().AppendText("Address1: 722 Moss Bay Blvd,  ");
			table[2, 1].AddParagraph().AppendText("KirklandWA USA");
			//Adds new Paragraph to the section.
			paragraph = section.AddParagraph();
			paragraph.AppendBreak(BreakType.PageBreak);
			//Adds text range.
			paragraph.AppendText("Product Overview");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);
			paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Justify;
			//Adds a new table into Word document.
			table = section.AddTable();
			//Specifies the total number of rows & columns.
			table.ResetCells(3, 2);
			//Accesses the instance of the cell  and adds the content into cell.
			//First row.
			imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Mountain-200.png"), FileMode.Open, FileAccess.ReadWrite);
			table[0, 0].AddParagraph().AppendPicture(imageStream);
			table[0, 1].AddParagraph().AppendText("Mountain-200");
			paragraph = table[0, 1].AddParagraph();
			paragraph.AppendText("Product No: ");
			paragraph.AppendField("Id", FieldType.FieldSequence);
			table[0, 1].AddParagraph().AppendText("Size: 38");
			table[0, 1].AddParagraph().AppendText("Weight: 25");
			table[0, 1].AddParagraph().AppendText("Price: $2,294.99");
			//Second row.
			table[1, 0].AddParagraph().AppendText("Mountain-300");
			paragraph = table[1, 0].AddParagraph();
			paragraph.AppendText("Product No: ");
			paragraph.AppendField("Id", FieldType.FieldSequence);
			table[1, 0].AddParagraph().AppendText("Size: 35");
			table[1, 0].AddParagraph().AppendText("Weight: 22");
			table[1, 0].AddParagraph().AppendText("Price: $1,079.99");
			imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Mountain-300.png"), FileMode.Open, FileAccess.ReadWrite);
			table[1, 1].AddParagraph().AppendPicture(imageStream);
			//Third row.
			imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Road-550.png"), FileMode.Open, FileAccess.ReadWrite);
			table[2, 0].AddParagraph().AppendPicture(imageStream);
			table[2, 1].AddParagraph().AppendText("Road-150");
			paragraph = table[2, 1].AddParagraph();
			paragraph.AppendText("Product No: ");
			paragraph.AppendField("Id", FieldType.FieldSequence);
			table[2, 1].AddParagraph().AppendText("Size: 44");
			table[2, 1].AddParagraph().AppendText("Weight: 14");
			table[2, 1].AddParagraph().AppendText("Price: $3,578.27");
			return document;
		}
    }
}
