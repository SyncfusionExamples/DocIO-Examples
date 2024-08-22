using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Fill_form_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    IWSection sec = document.LastSection;
                    InlineContentControl inlineCC;
                    InlineContentControl dropDownCC;
                    WTable table1 = sec.Tables[1] as WTable;
                    WTableRow row1 = table1.Rows[1];

                    #region General Information
                    //Fill the name.
                    WParagraph cellPara1 = row1.Cells[0].ChildEntities[1] as WParagraph;
                    inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    WTextRange text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "Steve Jobs";
                    inlineCC.ParagraphItems.Add(text);
                    //Fill the date of birth.
                    cellPara1 = row1.Cells[0].ChildEntities[3] as WParagraph;
                    inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "06/01/1994";
                    inlineCC.ParagraphItems.Add(text);
                    //Fill the address.
                    cellPara1 = row1.Cells[0].ChildEntities[5] as WParagraph;
                    inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "2501 Aerial Center Parkway.";
                    inlineCC.ParagraphItems.Add(text);
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "Morrisville, NC 27560.";
                    inlineCC.ParagraphItems.Add(text);
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "USA.";
                    inlineCC.ParagraphItems.Add(text);
                    //Fill the phone no.
                    cellPara1 = row1.Cells[0].ChildEntities[7] as WParagraph;
                    inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "+1 919.481.1974";
                    inlineCC.ParagraphItems.Add(text);
                    //Fill the email id.
                    cellPara1 = row1.Cells[0].ChildEntities[9] as WParagraph;
                    inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
                    text.Text = "steve123@email.com";
                    inlineCC.ParagraphItems.Add(text);
                    #endregion

                    #region Educational Information
                    table1 = sec.Tables[2] as WTable;
                    row1 = table1.Rows[1];
                    //Fill the education type.
                    cellPara1 = row1.Cells[0].ChildEntities[1] as WParagraph;
                    dropDownCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
                    text.Text = dropDownCC.ContentControlProperties.ContentControlListItems[1].DisplayText;
                    dropDownCC.ParagraphItems.Add(text);
                    //Fill the university.
                    cellPara1 = row1.Cells[0].ChildEntities[3] as WParagraph;
                    inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
                    text.Text = "Michigan University";
                    inlineCC.ParagraphItems.Add(text);
                    //Fill the C# experience level.
                    cellPara1 = row1.Cells[0].ChildEntities[7] as WParagraph;
                    dropDownCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
                    text.Text = dropDownCC.ContentControlProperties.ContentControlListItems[2].DisplayText;
                    dropDownCC.ParagraphItems.Add(text);
                    //Fill the VB experience level.
                    cellPara1 = row1.Cells[0].ChildEntities[9] as WParagraph;
                    dropDownCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
                    text = new WTextRange(document);
                    text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
                    text.Text = dropDownCC.ContentControlProperties.ContentControlListItems[1].DisplayText;
                    dropDownCC.ParagraphItems.Add(text);
                    #endregion
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
