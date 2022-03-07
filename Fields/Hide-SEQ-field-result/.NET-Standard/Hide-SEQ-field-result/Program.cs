using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Hide_SEQ_field_result
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a Word document.
            using (WordDocument document = CreateDocument())
            {
				//Accesses sequence field in the document.
				WTable table = document.LastSection.Body.ChildEntities[1] as WTable;
				WSeqField field = ((table[2, 1].ChildEntities[0] as WParagraph).ChildEntities[0] as WSeqField);
				//Enables a flag to to hide the sequence field result .
				field.HideResult = true;
				//Accesses sequence field in the document.
				field = ((table[4, 1].ChildEntities[0] as WParagraph).ChildEntities[0] as WSeqField);
				//Enables a flag to hide the sequence field result.
				field.HideResult = true;
				//Updates the document fields
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
			//Creates a new Word document.
			WordDocument document = new WordDocument();
			//Adds a new section to the document.
			IWSection section = document.AddSection();
			//Adds a paragraph to the section.
			IWParagraph paragraph = section.AddParagraph();
			paragraph.AppendText("Syncfusion Product Details");
			paragraph.ApplyStyle(BuiltinStyle.Heading1);
			//Adds a new table .
			IWTable table = section.AddTable();
			//Specifies the total number of rows & columns.
			table.ResetCells(6, 4);
			//Accesses the instance of the cell and add the content into cell.
			//First row.
			table[0, 0].AddParagraph().AppendText("S.No");
			table[0, 1].AddParagraph().AppendText("Platform Id");
			table[0, 2].AddParagraph().AppendText("Platform");
			table[0, 3].AddParagraph().AppendText("Status ");
			table[1, 0].AddParagraph().AppendText("1.");
			//Second row.
			table[1, 1].AddParagraph().AppendField("PlatformCount", FieldType.FieldSequence);
			table[1, 2].AddParagraph().AppendText("ASP.NET Core");
			table[1, 3].AddParagraph().AppendText("Live");
			//Third row.
			table[2, 0].AddParagraph().AppendText("2.");
			table[2, 1].AddParagraph().AppendField("PlatformCount", FieldType.FieldSequence);
			table[2, 2].AddParagraph().AppendText("LightSwitch");
			table[2, 3].AddParagraph().AppendText("Retired");
			//Fourth row.
			table[3, 0].AddParagraph().AppendText("3.");
			table[3, 1].AddParagraph().AppendField("PlatformCount", FieldType.FieldSequence);
			table[3, 2].AddParagraph().AppendText("ASP.NET MVC");
			table[3, 3].AddParagraph().AppendText("Live");
			//Fifth row.
			table[4, 0].AddParagraph().AppendText("4.");
			table[4, 1].AddParagraph().AppendField("PlatformCount", FieldType.FieldSequence);
			table[4, 2].AddParagraph().AppendText("Silverlight ");
			table[4, 3].AddParagraph().AppendText("Retired");
			//Sixth row.
			table[5, 0].AddParagraph().AppendText("5.");
			table[5, 1].AddParagraph().AppendField("PlatformCount", FieldType.FieldSequence);
			table[5, 2].AddParagraph().AppendText("Blazor");
			table[5, 3].AddParagraph().AppendText("Live");
			section.AddParagraph().AppendText("Total No. of Platforms : 5");
			return document;
		}
    }
}
