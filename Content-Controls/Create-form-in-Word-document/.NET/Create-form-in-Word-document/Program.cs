using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_form_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adding a new section to the document.
                IWSection section = document.AddSection();
                //Adding a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();

                #region Document formatting
                //Sets background color for document.
                document.Background.Gradient.Color1 = Syncfusion.Drawing.Color.FromArgb(232, 232, 232);
                document.Background.Gradient.Color2 = Syncfusion.Drawing.Color.FromArgb(255, 255, 255);
                document.Background.Type = BackgroundType.Gradient;
                document.Background.Gradient.ShadingStyle = GradientShadingStyle.Horizontal;
                document.Background.Gradient.ShadingVariant = GradientShadingVariant.ShadingDown;
                //Sets page size for document.
                section.PageSetup.Margins.All = 30f;
                section.PageSetup.PageSize = new Syncfusion.Drawing.SizeF(600, 600f);
                #endregion

                #region Title Section
                //Adds a new table to the section.
                IWTable table = section.Body.AddTable();
                table.ResetCells(1, 2);
                //Gets the table first row.
                WTableRow row = table.Rows[0];
                row.Height = 25f;
                //Adds a new paragraph to the cell.
                IWParagraph cellPara = row.Cells[0].AddParagraph();
                //Appends new picture.
                IWPicture pic = cellPara.AppendPicture(new FileStream(Path.GetFullPath(@"../../../image.jpg"), FileMode.Open, FileAccess.Read));
                pic.Height = 80;
                pic.Width = 180;
                //Adds a new paragraph to the next cell.
                cellPara = row.Cells[1].AddParagraph();
                row.Cells[1].CellFormat.VerticalAlignment = Syncfusion.DocIO.DLS.VerticalAlignment.Middle;
                row.Cells[1].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(173, 215, 255);
                //Appends the text.
                IWTextRange txt = cellPara.AppendText("Job Application Form");
                cellPara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
                //Sets the formats.
                txt.CharacterFormat.Bold = true;
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 18f;
                //Sets the widths and border types.
                row.Cells[0].Width = 200;
                row.Cells[1].Width = 300;
                row.Cells[1].CellFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Hairline;
                #endregion
                //Adds a new paragraph.
                section.AddParagraph();

                #region General Information
                //Adds a new table.
                table = section.Body.AddTable();
                table.ResetCells(2, 1);
                row = table.Rows[0];
                row.Height = 20;
                row.Cells[0].Width = 500;
                //Adds a new paragraph.
                cellPara = row.Cells[0].AddParagraph();
                //Sets a border type, color and background for cell.
                row.Cells[0].CellFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                row.Cells[0].CellFormat.Borders.Color = Syncfusion.Drawing.Color.FromArgb(155, 205, 255);
                row.Cells[0].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(198, 227, 255);
                row.Cells[0].CellFormat.VerticalAlignment = Syncfusion.DocIO.DLS.VerticalAlignment.Middle;
                txt = cellPara.AppendText(" General Information");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.Bold = true;
                txt.CharacterFormat.FontSize = 11f;
                row = table.Rows[1];
                cellPara = row.Cells[0].AddParagraph();
                //Sets a width, border type, color and background for cell.
                row.Cells[0].Width = 500;
                row.Cells[0].CellFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Hairline;
                row.Cells[0].CellFormat.Borders.Color = Syncfusion.Drawing.Color.FromArgb(155, 205, 255);
                row.Cells[0].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(222, 239, 255);
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n Full Name:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                InlineContentControl txtField = cellPara.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                txtField.ContentControlProperties.Title = "Text";
                //Sets formatting options for text present insider a content control.
                txtField.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                txtField.BreakCharacterFormat.FontName = "Arial";
                txtField.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n Birth Date:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                txtField = cellPara.AppendInlineContentControl(ContentControlType.Date) as InlineContentControl;
                txtField.ContentControlProperties.Title = "Date";
                //Sets the date display format
                txtField.ContentControlProperties.DateDisplayFormat = "M/d/yyyy";
                //Sets formatting options for text present insider a content control.
                txtField.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                txtField.BreakCharacterFormat.FontName = "Arial";
                txtField.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n Address:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                txtField = cellPara.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                txtField.ContentControlProperties.Title = "Text";
                //Sets multiline property to true to get the multiple line input of Address.
                txtField.ContentControlProperties.Multiline = true;
                //Sets formatting options for text present insider a content control.
                txtField.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                txtField.BreakCharacterFormat.FontName = "Arial";
                txtField.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n Phone:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                txtField = cellPara.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                txtField.ContentControlProperties.Title = "Text";
                //Sets formatting options for text present insider a content control.
                txtField.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                txtField.BreakCharacterFormat.FontName = "Arial";
                txtField.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n Email:\t\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                txtField = cellPara.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                txtField.ContentControlProperties.Title = "Text";
                //Sets formatting options for text present insider a content control.
                txtField.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                txtField.BreakCharacterFormat.FontName = "Arial";
                txtField.BreakCharacterFormat.FontSize = 11f;
                cellPara.AppendText("\n");
                #endregion
                section.AddParagraph();

                #region Educational Qualification
                //Adds a new table to the section.
                table = section.Body.AddTable();
                table.ResetCells(2, 1);
                row = table.Rows[0];
                row.Height = 20;
                //Sets width, border type, color, background and vertical alignment for cell.
                row.Cells[0].Width = 500;
                row.Cells[0].CellFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Thick;
                row.Cells[0].CellFormat.Borders.Color = Syncfusion.Drawing.Color.FromArgb(155, 205, 255);
                row.Cells[0].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(198, 227, 255);
                row.Cells[0].CellFormat.VerticalAlignment = Syncfusion.DocIO.DLS.VerticalAlignment.Middle;
                cellPara = row.Cells[0].AddParagraph();
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText(" Educational Qualification");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.Bold = true;
                txt.CharacterFormat.FontSize = 11f;
                row = table.Rows[1];
                //Sets width, border type, color, and background for cell.
                row.Cells[0].Width = 500;
                row.Cells[0].CellFormat.Borders.BorderType = Syncfusion.DocIO.DLS.BorderStyle.Hairline;
                row.Cells[0].CellFormat.Borders.Color = Syncfusion.Drawing.Color.FromArgb(155, 205, 255);
                row.Cells[0].CellFormat.BackColor = Syncfusion.Drawing.Color.FromArgb(222, 239, 255);
                cellPara = row.Cells[0].AddParagraph();
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n Type:\t\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                InlineContentControl dropdown = cellPara.AppendInlineContentControl(ContentControlType.DropDownList) as InlineContentControl;
                WTextRange textRange = new WTextRange(document);
                textRange.Text = "Choose an item from drop down list";
                dropdown.ParagraphItems.Add(textRange);
                //Creates an item for dropdown list.
                ContentControlListItem item = new ContentControlListItem();
                //Sets the text to be displayed as list item.
                item.DisplayText = "Higher";
                //Sets the value to the list item.
                item.Value = "1";
                //Adds item to the dropdown list.
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Vocational";
                item.Value = "2";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Universal";
                item.Value = "3";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                dropdown.ContentControlProperties.Title = "Drop-Down";
                //Sets formatting options for text present insider a content control.
                dropdown.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                dropdown.BreakCharacterFormat.FontName = "Arial";
                dropdown.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n Institution:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a new inline content control to enter the value.
                txtField = cellPara.AppendInlineContentControl(ContentControlType.Text) as InlineContentControl;
                //Sets formatting options for text present insider a content control.
                txtField.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                txtField.BreakCharacterFormat.FontName = "Arial";
                txtField.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n Programming Languages:");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n\t C#:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 9f;
                //Appends a new inline content control to enter the value.
                dropdown = cellPara.AppendInlineContentControl(ContentControlType.DropDownList) as InlineContentControl;
                textRange = new WTextRange(document);
                textRange.Text = "Choose an item from drop down list";
                dropdown.ParagraphItems.Add(textRange);
                //Creates an item for dropdown list.
                item = new ContentControlListItem();
                //Sets the text to be displayed as list item.
                item.DisplayText = "Perfect";
                //Sets the value to the list item.
                item.Value = "1";
                //Adds item to the dropdown list.
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Good";
                item.Value = "2";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Excellent";
                item.Value = "3";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                //Sets formatting options for text present insider a content control.
                dropdown.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                dropdown.BreakCharacterFormat.FontName = "Arial";
                dropdown.BreakCharacterFormat.FontSize = 11f;
                //Appends a text to paragraph of cell.
                txt = cellPara.AppendText("\n\n\t VB:\t\t\t\t");
                txt.CharacterFormat.FontName = "Arial";
                txt.CharacterFormat.FontSize = 9f;
                //Appends a new inline content control to enter the value.
                dropdown = cellPara.AppendInlineContentControl(ContentControlType.DropDownList) as InlineContentControl;
                textRange = new WTextRange(document);
                textRange.Text = "Choose an item from drop down list";
                dropdown.ParagraphItems.Add(textRange);
                //Creates an item for dropdown list.
                item = new ContentControlListItem();
                //Sets the text to be displayed as list item.
                item.DisplayText = "Perfect";
                //Sets the value to the list item.
                item.Value = "1";
                //Adds item to the dropdown list.
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Good";
                item.Value = "2";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                item = new ContentControlListItem();
                item.DisplayText = "Excellent";
                item.Value = "3";
                dropdown.ContentControlProperties.ContentControlListItems.Add(item);
                dropdown.ContentControlProperties.Title = "Drop-Down";
                //Sets formatting options for text present insider a content control
                dropdown.BreakCharacterFormat.TextColor = Syncfusion.Drawing.Color.MidnightBlue;
                dropdown.BreakCharacterFormat.FontName = "Arial";
                dropdown.BreakCharacterFormat.FontSize = 11f;
                #endregion
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
