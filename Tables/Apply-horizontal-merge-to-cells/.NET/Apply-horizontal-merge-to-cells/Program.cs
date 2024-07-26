using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Apply_horizontal_merge_to_cells
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                section.AddParagraph().AppendText("Horizontal merging of Table cells");
                IWTable table = section.AddTable();
                table.ResetCells(5, 5);
                //Specifies the horizontal merge from second cell to fifth cell in third row.
                table.ApplyHorizontalMerge(2, 1, 4);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
