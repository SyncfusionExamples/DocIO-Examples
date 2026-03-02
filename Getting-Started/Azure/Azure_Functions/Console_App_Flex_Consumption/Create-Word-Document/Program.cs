using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

class Program
{
    static async Task Main()
    {
        try
        {
            Console.Write("Please enter your Azure Function URL: ");
            string url = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(url)) return;
            // Create a new HttpClient instance for sending HTTP requests
            using var http = new HttpClient();
            using var content = new StringContent(string.Empty);
            using var res = await http.PostAsync(url, content);
            // Read the response body as a byte array
            var resBytes = await res.Content.ReadAsByteArrayAsync();
            // Extract the media type from the response headers
            string mediaType = res.Content.Headers.ContentType?.MediaType ?? string.Empty;
            // Decide the output file path the response is an docx or txt   
            string outputPath = mediaType.Contains("word", StringComparison.OrdinalIgnoreCase)
                || mediaType.Contains("officedocument", StringComparison.OrdinalIgnoreCase)
                || mediaType.Equals("application/vnd.openxmlformats-officedocument.wordprocessingml.document", StringComparison.OrdinalIgnoreCase)
                ? Path.GetFullPath("../../../Output/Output.docx")
                : Path.GetFullPath("../../../Output/function-error.txt");
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