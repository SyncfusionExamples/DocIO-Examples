using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Drawing;

namespace Add_watermark_specificpage
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document file for reading
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the document into the WordDocument object
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Retrieve all sections in the document
                    List<Entity> sections = document.FindAllItemsByProperty(EntityType.Section, null, null);

                    // Add "Syncfusion" watermark to the first section
                    AddWatermarkToPage(sections[0] as WSection, "Adventures");

                    // Add "Draft" watermark to the second section
                    AddWatermarkToPage(sections[1] as WSection, "Pandas");

                    // Save the modified document to a new file
                    using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(docStream1, FormatType.Docx);
                    }
                }
            }
        }

        // Method to add a watermark in the document
        static void AddWatermarkToPage(IWSection section, string watermarkText)
        {
            // Access the body of the specified section
            WTextBody textBody = section.Body;
            // Adds a block content control (RichText) to the section
            BlockContentControl blockContentControl = textBody.AddBlockContentControl(ContentControlType.RichText) as BlockContentControl;

            // Adds a new paragraph inside the block content control
            WParagraph paragraph = blockContentControl.TextBody.AddParagraph() as WParagraph;
            // Create a  text box to hold the watermark text
            WTextBox watermarkTextBox = paragraph.AppendTextBox(494, 164) as WTextBox;
            // Center-align the text box horizontally
            watermarkTextBox.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            // Center-align the text box vertically
            watermarkTextBox.TextBoxFormat.VerticalAlignment = ShapeVerticalAlignment.Center;
            // Remove the border line of the text box
            watermarkTextBox.TextBoxFormat.NoLine = true;
            // Set rotation angle for the watermark text box
            watermarkTextBox.TextBoxFormat.Rotation = 315;
            // Allow overlapping of the text box with other elements
            watermarkTextBox.TextBoxFormat.AllowOverlap = true;
            // Align the text box relative to the page margins
            watermarkTextBox.TextBoxFormat.HorizontalOrigin = HorizontalOrigin.Margin;
            watermarkTextBox.TextBoxFormat.VerticalOrigin = VerticalOrigin.Margin;
            // Set text wrapping style to behind, so the watermark does not interfere with content
            watermarkTextBox.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.Behind;

            // Add another paragraph inside the text box to contain the watermark text
            IWParagraph watermarkParagraph = watermarkTextBox.TextBoxBody.AddParagraph();
            // Append the watermark text to the paragraph and set the font size and color
            IWTextRange textRange = watermarkParagraph.AppendText(watermarkText);
            // Set a large font size for the watermark text
            textRange.CharacterFormat.FontSize = 100;
            // Set a light gray color for the watermark text
            textRange.CharacterFormat.TextColor = Color.FromArgb(255, 192, 192, 192);
        }
    }
}
