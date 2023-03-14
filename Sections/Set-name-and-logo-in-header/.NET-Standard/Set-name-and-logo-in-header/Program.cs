using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_name_and_logo_in_header
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2VVhkQlFadV5JXGFWfVJpTGpQdk5xdV9DaVZUTWY/P1ZhSXxRd0djXn5ZcXVQRWVfVEA=");
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                section.PageSetup.DifferentFirstPage = true;
                //Add a table to the header.
                IWTable table = section.HeadersFooters.FirstPageHeader.AddTable();
                table.ResetCells(1, 2);
                table.TableFormat.Borders.BorderType = BorderStyle.Single;
                //Add paragraph to the first cell.
                IWParagraph paragraph = table[0, 0].AddParagraph();
                //Set the alignment to Left.
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
                IWTextRange textRange = paragraph.AppendText("Adventure Works Cycles");
                //Add paragraph to the second cell.
                paragraph = table[0, 1].AddParagraph();
                //Add image to  the paragraph.
                FileStream imageStream = new FileStream(@"../../../Logo.jpg", FileMode.Open, FileAccess.ReadWrite);
                IWPicture picture = paragraph.AppendPicture(imageStream);
                //Set the alignment to right.
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                picture.Width = 120;
                picture.Height = 80;
                //Create file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
