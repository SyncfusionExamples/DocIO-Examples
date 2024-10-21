using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Net;

ResizeImageUsingAltText();

static void ResizeImageUsingAltText()
{
    using (WordDocument document = new WordDocument())
    {
        // Register the event to customize images while importing Markdown.
        document.MdImportSettings.ImageNodeVisited += MdImportSettings_ImageNodeVisited;
        // Open the input Markdown file for reading.
        using (FileStream inputFileStream = new FileStream(@"Data/Input.md", FileMode.Open, FileAccess.Read))
        {
            document.Open(inputFileStream, FormatType.Markdown);
            #region ImageResize
            // Find all images with the alternative text "File" and resize them to 300x300.
            List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, "AlternativeText", "File");
            foreach (WPicture picture in pictures)
            {
                picture.Height = 284;
                picture.Width = 442;
            }
            #endregion
            // Save the modified document.
            using (FileStream outputFileStream = new FileStream(@"Output/Result.docx", FileMode.Create, FileAccess.Write))
            {
                document.Save(outputFileStream, FormatType.Docx);
            } 
        }
    }
}

static void MdImportSettings_ImageNodeVisited(object sender, Syncfusion.Office.Markdown.MdImageNodeVisitedEventArgs args)
{
    // Set the image stream based on the image name from the Markdown input.
    if (args.Uri == "Image_1.png")
        args.ImageStream = new FileStream(@"Data/Image_1.png", FileMode.Open);
    else if (args.Uri == "Image_2.png")
        args.ImageStream = new FileStream(@"Data/Image_2.png", FileMode.Open);
    // If the image is from a URL, download and set it as a stream.
    else if (args.Uri.StartsWith("https://"))
    {
        // Create a WebClient instance.
        WebClient client = new WebClient();
        // Download the image as byte data.
        byte[] image = client.DownloadData(args.Uri);
        // Convert byte data to a memory stream.
        Stream stream = new MemoryStream(image);
        // Set the stream for the image in Markdown.
        args.ImageStream = stream; 
    }
}
