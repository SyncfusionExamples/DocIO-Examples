using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

class Program
{
    static async Task Main()
    {  
        Console.Write("Please enter your Azure Functions URL : ");
        // Read the URL entered by the user and trim whitespace
        string url = Console.ReadLine()?.Trim();
        // If no URL was entered, exit the program
        if (string.IsNullOrEmpty(url)) return;
        // Create a new HttpClient instance for sending requests
        using var http = new HttpClient();
        // Read all bytes from the input Word document file
        var bytes = await File.ReadAllBytesAsync(@"Data/Input.docx");
        // Create HTTP content from the document bytes
        using var content = new ByteArrayContent(bytes);
        // Set the content type header to application/octet-stream (binary data)
        content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
        // Send a POST request to the Azure Function with the document content
        using var res = await http.PostAsync(url, content);
        // Read the response content as a byte array
        var resBytes = await res.Content.ReadAsByteArrayAsync();
        // Get the media type (e.g., application/pdf or text/plain) from the response headers
        string mediaType = res.Content.Headers.ContentType?.MediaType ?? string.Empty;
        string outFile = mediaType.Contains("pdf", StringComparison.OrdinalIgnoreCase)
            ? Path.GetFullPath(@"../../../Output/Output.pdf")
            : Path.GetFullPath(@"../../../Output/function-error.txt");
        // Write the response bytes to the chosen output file
        await File.WriteAllBytesAsync(outFile, resBytes);
        Console.WriteLine($"Saved: {outFile} ");        
    }
}