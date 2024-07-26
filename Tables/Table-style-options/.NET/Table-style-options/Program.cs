using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Table_style_options
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    WSection section = document.Sections[0];
                    WTable table = section.Tables[0] as WTable;
                    //Applies "LightShading" built-in style to table.
                    table.ApplyStyle(BuiltinTableStyle.LightShading);
                    //Enables special formatting for banded columns of the table.
                    table.ApplyStyleForBandedColumns = true;
                    //Enables special formatting for banded rows of the table.
                    table.ApplyStyleForBandedRows = true;
                    //Disables special formatting for first column of the table.
                    table.ApplyStyleForFirstColumn = false;
                    //Enables special formatting for header row of the table.
                    table.ApplyStyleForHeaderRow = true;
                    //Enables special formatting for last column of the table.
                    table.ApplyStyleForLastColumn = true;
                    //Disables special formatting for last row of the table.
                    table.ApplyStyleForLastRow = false;
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
}
