using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Align_text_within_a_table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
					//Gets the text body of first section
					WTextBody textBody = document.Sections[0].Body;
					//Iterates the cells within a table and align text for each cell.
					AlignCellContentForTextBody(textBody, HorizontalAlignment.Center, VerticalAlignment.Middle);
					//Creates file stream.
					using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

		/// <summary>
		/// Iterates the child items of textbody.
		/// </summary>
		private static void AlignCellContentForTextBody(WTextBody textBody, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
		{
			for (int i = 0; i < textBody.ChildEntities.Count; i++)
			{
				//IEntity is the basic unit in DocIO DOM. 
				//Accesses the body items as IEntity
				IEntity bodyItemEntity = textBody.ChildEntities[i];
				//A Text body has 3 types of elements - Paragraph, Table and Block Content Control.
				//Decides the element type by using EntityType.
				switch (bodyItemEntity.EntityType)
				{
					case EntityType.Paragraph:
						WParagraph paragraph = bodyItemEntity as WParagraph;
						//Sets horizontal alignment for paragraph.
						paragraph.ParagraphFormat.HorizontalAlignment = horizontalAlignment;
						break;
					case EntityType.Table:
						//Table is a collection of rows and cells.
						//Iterates through table's DOM and set horizontal alignment.
						AlignCellContentForTable(bodyItemEntity as WTable, horizontalAlignment, verticalAlignment);
						break;
					case EntityType.BlockContentControl:
						BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
						//Iterates to the body items of Block Content Control and set horizontal alignment.
						AlignCellContentForTextBody(blockContentControl.TextBody, horizontalAlignment, verticalAlignment);
						break;
				}
			}
		}

		/// <summary>
		/// Iterates the child items of table.s
		/// </summary>
		private static void AlignCellContentForTable(WTable table, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
		{
			//Iterates the row collection in a table.
			foreach (WTableRow row in table.Rows)
			{
				//Iterates the cell collection in a table row.
				foreach (WTableCell cell in row.Cells)
				{
					//Sets vertical alignment to the cell.
					cell.CellFormat.VerticalAlignment = verticalAlignment;
					//Iterate items in cell and set horizontal alignment.
					AlignCellContentForTextBody(cell, horizontalAlignment, verticalAlignment);
				}
			}
		}
	}
}
