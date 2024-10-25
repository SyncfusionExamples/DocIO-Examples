using System.Net;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");

// Request URLs for header, footer, and main body content.
Console.WriteLine("Please enter the URL for the header content:");
string headerHtmlUrl = Console.ReadLine(); 
Console.WriteLine("Please enter the URL for the footer content:");
string footerHtmlUrl = Console.ReadLine(); 
Console.WriteLine("Please enter the URL for the main body content:");
string bodyHtmlUrl = Console.ReadLine();
// Retrieve HTML content from the specified URLs.
string headerContent = GetHtmlContent(headerHtmlUrl);
string footerContent = GetHtmlContent(footerHtmlUrl);
string mainContent = GetHtmlContent(bodyHtmlUrl);
// Create a new Word document instance.
using (WordDocument document = new WordDocument())
{
    // Add a new section to the document.
    WSection section = document.AddSection() as WSection;
    // Append the main content HTML to the paragraph.
    WParagraph paragraph = section.AddParagraph() as WParagraph;
    paragraph.AppendHTML(mainContent);
    // Append the header content HTML to the header paragraph.
    paragraph = section.HeadersFooters.OddHeader.AddParagraph() as WParagraph;
    paragraph.AppendHTML(headerContent);
    // Append the footer content HTML to the footer paragraph.
    paragraph = section.HeadersFooters.OddFooter.AddParagraph() as WParagraph;
    paragraph.AppendHTML(footerContent); 
    // Save the modified document.
    using (FileStream outputStream = new FileStream("Output/Output.docx", FileMode.Create, FileAccess.Write))
    {
        document.Save(outputStream, FormatType.Docx); // Save the document in DOCX format.
    }
}

/// <summary>
/// Fetches the HTML content from a given URL by sending a GET request and reading the server's response stream.
/// </summary>
string GetHtmlContent(string url)
{
    // Create a web request to the specified URL.
    WebRequest myRequest = WebRequest.Create(url);
    // Set the request method to GET to fetch data from the URL.
    myRequest.Method = "GET";
    // Get the response from the web server.
    WebResponse myResponse = myRequest.GetResponse();
    // Read the response stream and return the HTML content as a string.
    using (StreamReader sr = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.UTF8))
    {
        // Read all content from the response stream.
        string result = sr.ReadToEnd();
        // Return the HTML content as a string.
        return result;
    }
}

