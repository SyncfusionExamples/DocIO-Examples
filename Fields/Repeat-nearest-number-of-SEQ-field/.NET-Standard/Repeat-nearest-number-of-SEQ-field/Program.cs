using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Repeat_nearest_number_of_SEQ_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a Word document.
            using (WordDocument document = CreateDocument())
            {
                //Accesses sequence field in the document.
                WSeqField field = (document.LastSection.HeadersFooters.Header.ChildEntities[0] as WParagraph).ChildEntities[1] as WSeqField;
                //Enables a flag to repeat the nearest number for sequence field.
                field.RepeatNearestNumber = true;
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

        private static WordDocument CreateDocument()
        {
			//Creates a new document.
			WordDocument document = new WordDocument();
			//Adds a new section to the document.
			IWSection section = document.AddSection();
			//Inserts the default page header.
			IWParagraph paragraph = section.HeadersFooters.OddHeader.AddParagraph();
			paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
			paragraph.AppendText("Total No. of Products: ");
			paragraph.AppendField("Product count", FieldType.FieldSequence);
			//Adds a paragraph to the section.
			paragraph = section.AddParagraph();
			IWTextRange textRange = paragraph.AppendText("Adventure Works Cycles");
			paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
			textRange.CharacterFormat.FontSize = 16;
			textRange.CharacterFormat.Bold = true;
			//Adds a paragraph to the section.
			section.AddParagraph().AppendText("Product Overview");
			document.LastParagraph.ApplyStyle(BuiltinStyle.Heading1);
			//Adds a new table into Word document.
			IWTable table = section.AddTable();
			//Specifies the total number of rows & columns.
			table.ResetCells(3, 2);
			//Accesses the instance of the cell  and adds the content into cell.
			//First row.
			FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Mountain-200.png"), FileMode.Open, FileAccess.ReadWrite);
			table[0, 0].AddParagraph().AppendPicture(imageStream);
			table[0, 1].AddParagraph().AppendText("Mountain-200");
			paragraph = table[0, 1].AddParagraph();
			paragraph.AppendText("Product No: ");
			paragraph.AppendField("Product count", FieldType.FieldSequence);
			table[0, 1].AddParagraph().AppendText("Size: 38");
			table[0, 1].AddParagraph().AppendText("Weight: 25");
			table[0, 1].AddParagraph().AppendText("Price: $2,294.99");
			//Second row.
			table[1, 0].AddParagraph().AppendText("Mountain-300");
			paragraph = table[1, 0].AddParagraph();
			paragraph.AppendText("Product No: ");
			paragraph.AppendField("Product count", FieldType.FieldSequence);
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
			paragraph.AppendField("Product count", FieldType.FieldSequence);
			table[2, 1].AddParagraph().AppendText("Size: 44");
			table[2, 1].AddParagraph().AppendText("Weight: 14");
			table[2, 1].AddParagraph().AppendText("Price: $3,578.27");
			return document;
		}
    }
}
