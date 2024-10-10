using System.Net;
using Syncfusion.DocIO.DLS;

// Request URLs from the user
Console.WriteLine("Please enter the URL for the Header content:");
string mainPageURL = Console.ReadLine();

Console.WriteLine("Please enter the URL for the Footer content:");
string footerURL = Console.ReadLine();

Console.WriteLine("Please enter the URL for the Body content:");
string headerURL = Console.ReadLine();

// Fetch the HTML content from the specified URLs
string mainPage = code(mainPageURL);
string header = code(headerURL);
string footer = code(footerURL);

// Create a new Word document.
WordDocument document = new WordDocument();
// Add a new section to the document.
WSection section = document.AddSection() as WSection;
// Add a new paragraph in the main section.
WParagraph paragraph = section.AddParagraph() as WParagraph;
// Append the main page HTML content to the paragraph.
paragraph.AppendHTML(mainPage);
// Add a new paragraph to the Header section.
paragraph = section.HeadersFooters.OddHeader.AddParagraph() as WParagraph;
// Append the header HTML content to the header paragraph.
paragraph.AppendHTML(header);
// Add a new paragraph to the Footer section.
paragraph = section.HeadersFooters.OddFooter.AddParagraph() as WParagraph;
// Append the footer HTML content to the footer paragraph.
paragraph.AppendHTML(footer);
// Specify the output file path for the generated Word document.
using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.docx"), FileMode.Create, FileAccess.Write))
{
    // Save the generated Word document to the specified file stream in DOCX format.
    document.Save(outputStream, Syncfusion.DocIO.FormatType.Docx);
}
// Close the document to release resources.
document.Close();

/// <summary>
/// Fetches the HTML content from the specified URL.
/// </summary>
static string code(string Url)
{
    // Create a web request for the given URL.
    WebRequest myRequest = WebRequest.Create(Url);
    myRequest.Method = "GET";

    // Get the response from the web request.
    WebResponse myResponse = myRequest.GetResponse();
    StreamReader sr = new StreamReader(myResponse.GetResponseStream(), System.Text.Encoding.UTF8);

    // Read the response stream and store the HTML content in a string.
    string result = sr.ReadToEnd();

    // Close the StreamReader and WebResponse to release resources.
    sr.Close();
    myResponse.Close();

    // Return the HTML content.
    return result;
}
