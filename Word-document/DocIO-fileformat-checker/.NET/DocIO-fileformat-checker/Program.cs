using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace DocIO_fileformat_checker
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get all files from the directory
            string[] files = Directory.GetFiles(Path.GetFullPath(@"Data/"));
            // Loop through each file in the directory
            foreach (string filePath in files)
            {
                using (FileStream inputStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    try
                    {
                        // Try to open the file using Syncfusion DocIO
                        using (WordDocument doc = new WordDocument(inputStream, FormatType.Automatic))
                        {
                            // Successfully opened the document
                            Console.WriteLine("Supported format" + filePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Check if the exception message matches the unsupported format message
                        if (ex.Message.Contains("This file format is not supported"))
                        {
                            // If the file format is not supported, print it to the console
                            Console.WriteLine("Unsupported format: " + filePath);
                        }
                        else
                        {
                            // If some other exception occurs, handle it
                            Console.WriteLine($"Error opening file {filePath}: {ex.Message}");
                        }
                    }
                }
            }
        }
    }
}
