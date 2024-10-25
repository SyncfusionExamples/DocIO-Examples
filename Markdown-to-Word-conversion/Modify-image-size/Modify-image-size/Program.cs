using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Net;

//Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");

using (WordDocument document = new WordDocument())
{
    // Register an event to customize images while importing Markdown.
    document.MdImportSettings.ImageNodeVisited += MdImportSettings_ImageNodeVisited;
    // Open the input Markdown file.
    using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Input.md"), FileMode.Open, FileAccess.Read))
    {
        // Load the Markdown file into the Word document.
        document.Open(inputFileStream, FormatType.Markdown);

        // Find all images with the alternative text "Mountain" and resize them.
        List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, "AlternativeText", "Mountain");
        foreach (WPicture picture in pictures)
        {
            picture.Height = 250;
            picture.Width = 250;
        }

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
    // Load a specific image from a local path if the image URI matches.
    if (args.Uri == "Road-550.png")
        args.ImageStream = new FileStream(Path.GetFullPath(@"Data/Road-550.png"), FileMode.Open);
    // If the image URI starts with "https://", download and set the image from the URL.
    else if (args.Uri.StartsWith("https://"))
    {
        // Initialize a WebClient instance for downloading.
        WebClient client = new WebClient();
        // Download the image data as a byte array.
        byte[] image = client.DownloadData(args.Uri);
        // Convert the byte array to a memory stream.
        Stream stream = new MemoryStream(image);
        // Set the downloaded stream as the image in the Markdown.
        args.ImageStream = stream;
    }
}
