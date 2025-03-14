using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    // Load the Word document from the file stream
    using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
    {
        // Add header and footer to the document
        AddHeaderFooter(document);
        // Save the modified document to an output file
        using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
        {
            document.Save(outputStream1, FormatType.Docx);
        }
    }
}

/// <summary>
/// Adds an image to the header paragraph.
/// </summary>
/// <param name="headerParagraph">The paragraph where the image will be added.</param>
static void GetHeaderContent(IWParagraph headerParagraph)
{
    // Open the image file stream
    FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Picture.png"), FileMode.Open, FileAccess.Read);

    // Append the image to the header paragraph
    IWPicture picture = headerParagraph.AppendPicture(imageStream);

    // Set image positioning and formatting properties
    picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
    picture.VerticalOrigin = VerticalOrigin.Margin;
    picture.VerticalPosition = -45; // Adjust vertical position
    picture.HorizontalOrigin = HorizontalOrigin.Column;
    picture.HorizontalPosition = 30f; // Adjust horizontal position
    picture.WidthScale = 50; // Scale image width
    picture.HeightScale = 40; // Scale image height
}

/// <summary>
/// Adds a styled footer with a background color and page number.
/// </summary>
static void GetFooterContent(IWParagraph footerParagraph)
{
    // Create a rectangle shape in the footer for background styling
    Shape rectangleShape = footerParagraph.AppendShape(AutoShapeType.Rectangle, 700, 50);
    // Set rectangle shape properties (position, alignment, wrapping, color)
    rectangleShape.WrapFormat.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
    rectangleShape.HorizontalAlignment = ShapeHorizontalAlignment.Left;
    rectangleShape.VerticalAlignment = ShapeVerticalAlignment.Bottom;
    rectangleShape.HorizontalOrigin = HorizontalOrigin.Page;
    rectangleShape.VerticalOrigin = VerticalOrigin.Page;
    rectangleShape.FillFormat.Color = Syncfusion.Drawing.Color.LightBlue; // Set background color
    // Add a paragraph inside the rectangle for text
    footerParagraph = rectangleShape.TextBody.AddParagraph();
    footerParagraph.AppendText("Adventure Works Cycles");
    footerParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
    // Add a right-aligned tab stop for page number formatting
    footerParagraph.ParagraphFormat.Tabs.AddTab(523f, TabJustification.Right, TabLeader.NoLeader);
    // Append page number fields
    footerParagraph.AppendText("\t"); // Insert tab space
    footerParagraph.AppendField("Page", FieldType.FieldPage); // Current page number
    footerParagraph.AppendText(" of ");
    footerParagraph.AppendField("NumPages", FieldType.FieldNumPages); // Total number of pages
}

/// <summary>
/// Adds headers and footers to all sections of the document.
/// </summary>
static void AddHeaderFooter(WordDocument document)
{
    for (int i = 0; i < document.Sections.Count; i++)
    {
        WSection section = document.Sections[i];
        IWParagraph headerParagraph = section.HeadersFooters.FirstPageHeader.AddParagraph();
        GetHeaderContent(headerParagraph);
        IWParagraph headerOddParagraph = section.HeadersFooters.OddHeader.AddParagraph();
        GetHeaderContent(headerOddParagraph);
        IWParagraph evenParagraph = section.HeadersFooters.EvenHeader.AddParagraph();
        GetHeaderContent(evenParagraph);
        IWParagraph footerParagraph = section.HeadersFooters.FirstPageFooter.AddParagraph();
        GetFooterContent(footerParagraph);
        IWParagraph footerOddParagraph = section.HeadersFooters.OddFooter.AddParagraph();
        GetFooterContent(footerOddParagraph);
        IWParagraph footerEvenParagraph = section.HeadersFooters.EvenFooter.AddParagraph();
        GetFooterContent(footerEvenParagraph);
    }
}
