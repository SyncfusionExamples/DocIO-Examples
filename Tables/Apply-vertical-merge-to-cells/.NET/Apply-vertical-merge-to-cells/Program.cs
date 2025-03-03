﻿using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Apply_vertical_merge_to_cells
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                section.AddParagraph().AppendText("Vertical merging of Table cells");
                IWTable table = section.AddTable();
                table.ResetCells(5, 5);
                // Specifies the vertical merge to the third cell, from second row to fifth row.
                table.ApplyVerticalMerge(2, 1, 4);
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
