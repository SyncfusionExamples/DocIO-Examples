using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

class Program
{
    static async Task Main()
    {
        try
        {
            Console.Write("Please enter your Azure Functions URL : ");
            string url = Console.ReadLine();
            if (string.IsNullOrEmpty(url)) return;
            // Create a new HttpClient instance for sending HTTP requests
            using var http = new HttpClient();
            // Read all bytes from the input Word document 
            byte[] bytes = await File.ReadAllBytesAsync(@"Data/Input.docx");
            // Wrap the file bytes into a ByteArrayContent object for HTTP transmission
            using var content = new ByteArrayContent(bytes);
            // Set the content type header to indicate binary data
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
            // Send a POST request to the provided Azure Functions URL with the file content
            using var res = await http.PostAsync(url, content);
            // Read the response body as a byte array
            var resBytes = await res.Content.ReadAsByteArrayAsync();
            // Extract the media type from the response headers
            string mediaType = res.Content.Headers.ContentType?.MediaType ?? string.Empty;
            // Decide the output file path the response is an image or txt         
            string outputPath = mediaType.Contains("word", StringComparison.OrdinalIgnoreCase)
                || mediaType.Contains("officedocument", StringComparison.OrdinalIgnoreCase)
                || mediaType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document", StringComparison.OrdinalIgnoreCase)
                ? Path.GetFullPath(@"../../../Output/Output.docx")
                : Path.GetFullPath(@"../../../function-error.txt");
            // Write the response bytes to the output file 
            await File.WriteAllBytesAsync(outputPath, resBytes);
            Console.WriteLine($"Saved: {outputPath}");
        }
        catch (Exception ex)
        {
            throw;
        }


    }
}