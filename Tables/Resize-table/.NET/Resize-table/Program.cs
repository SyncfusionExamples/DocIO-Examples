using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Resize_table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates an instance of WordDocument class (Empty Word Document).
                using (WordDocument document = new WordDocument())
                {
                    //Opens an existing Word document into DocIO instance.
                    document.Open(fileStreamPath, FormatType.Docx);
                    //Accesses the instance of the first section in the Word document.
                    WSection section = document.Sections[0];
                    //Accesses the instance of the first table in the section.
                    WTable table = section.Tables[0] as WTable;
                    //Resizes the table to fit the contents respect to the contents.
                    table.AutoFit(AutoFitType.FitToContent);
                    //Accesses the instance of the second table in the section.
                    table = section.Tables[1] as WTable;
                    //Resizes the table to fit the contents respect to window/page width.
                    table.AutoFit(AutoFitType.FitToWindow);
                    //Accesses the instance of the third table in the section.
                    table = section.Tables[2] as WTable;
                    //Resizes the table to fit the contents respect to fixed column width.
                    table.AutoFit(AutoFitType.FixedColumnWidth);
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
}
