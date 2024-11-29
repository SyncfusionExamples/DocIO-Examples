using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_simple_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section and a paragraph to the document. 
                document.EnsureMinimal();
                //Add a new table to the Word document. 
                IWTable table = document.Sections[0].AddTable();
                //Specify the total number of rows and columns. 
                table.ResetCells(3, 2);
                //Access each table cell and append text. 
                table[0, 0].AddParagraph().AppendText("Item");
                table[0, 1].AddParagraph().AppendText("Price($)");
                table[1, 0].AddParagraph().AppendText("Apple");
                table[1, 1].AddParagraph().AppendText("50");
                table[2, 0].AddParagraph().AppendText("Orange");
                table[2, 1].AddParagraph().AppendText("30");
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
