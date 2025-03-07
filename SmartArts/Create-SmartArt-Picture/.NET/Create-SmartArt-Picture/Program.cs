using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
WTextRange textRange1 = paragraph.AppendText("Employee Report") as WTextRange;
textRange1.CharacterFormat.FontSize = 28f;
paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
// Add SmartArt with "Picture Strips" layout
WSmartArt pictureStripsSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.PictureStrips, 432, 252);
// Add the "Employee1" phase node
IOfficeSmartArtNode employeeNode1 = pictureStripsSmartArt.Nodes[0];
employeeNode1.TextBody.Text = "Nancy Davolio";
employeeNode1.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 25f;
AddPicture(employeeNode1, Path.GetFullPath(@"Images/Nancy Davolio.png"));
// Add the "Employee2" phase node
IOfficeSmartArtNode employeeNode2 = pictureStripsSmartArt.Nodes[1];
employeeNode2.TextBody.Text = "Andrew Fuller";
employeeNode2.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 25f;
AddPicture(employeeNode2, Path.GetFullPath(@"Images/Andrew Fuller.png"));
// Add the "Employee3" phase node
IOfficeSmartArtNode employeeNode3 = pictureStripsSmartArt.Nodes[2];
employeeNode3.TextBody.Text = "Janet Leverling";
employeeNode3.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 25f;
AddPicture(employeeNode3, Path.GetFullPath(@"Images/Janet Leverling.png"));
//Creates file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Saves the Word document to file stream.
    document.Save(outputFileStream, FormatType.Docx);
}

/// <summary>
/// Loads an image from the specified file path and assigns it to the given SmartArt node.
/// </summary>
void AddPicture(IOfficeSmartArtNode node, string imagePath)
{
    // Load the image and assign it to the SmartArt node
    FileStream pictureStream = new FileStream(imagePath, FileMode.Open);
    MemoryStream memoryStream = new MemoryStream();
    pictureStream.CopyTo(memoryStream);

    //Convert the memory stream into a byte array
    byte[] picByte = memoryStream.ToArray();
    node.Shapes[1].Fill.FillType = OfficeShapeFillType.Picture;
    node.Shapes[1].Fill.PictureFill.ImageBytes = picByte;
    //Dispose the image stream.
    pictureStream.Dispose();
    memoryStream.Dispose();
}