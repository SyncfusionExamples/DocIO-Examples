using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Net;

// Create a Word document instance.
using (WordDocument document = new WordDocument())
{
    // Hook the event to customize the image while importing Markdown.
    document.MdImportSettings.ImageNodeVisited += MdImportSettings_ImageNodeVisited;
    using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Input.md"), FileMode.Open, FileAccess.Read))
    {
        // Open the input Markdown file.
        document.Open(inputFileStream, FormatType.Markdown);
        // Find all images with the alternative text "Mountain" and resize them.
        WPicture picture = document.FindItemByProperty(EntityType.Picture, "AlternativeText", "Mountain") as WPicture;
        picture.Height = 250;
        picture.Width = 250;
        // Save the modified document.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
        {
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Customizes image loading from local and remote sources during Markdown import.
/// </summary>
static void MdImportSettings_ImageNodeVisited(object sender, Syncfusion.Office.Markdown.MdImageNodeVisitedEventArgs args)
{
    // Set the image stream based on the image name from the input Markdown.
    if (args.Uri == "Road-550.png")
        args.ImageStream = new FileStream(Path.GetFullPath(@"Data/Road-550.png"), FileMode.Open);
    // Retrieve the image from the website and use it.
    else if (args.Uri.StartsWith("https://"))
    {
        WebClient client = new WebClient();
        // Download the image as a stream.
        byte[] image = client.DownloadData(args.Uri);
        Stream stream = new MemoryStream(image);
        // Set the retrieved image from the input Markdown.
        args.ImageStream = stream;
    }
}
