using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs;
using Microsoft.AspNetCore.Mvc;
using Open_Word_document.Models;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;
using Syncfusion.DocIO.DLS;
using System.IO;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using System.IO.Pipes;

namespace Open_Word_document.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        public async Task<IActionResult> EditDocument()
        {
            //Your Azure Storage Account connection string
            string connectionString = "Your_connection_string";

            //Name of the Azure Blob Storage container
            string containerName = "Your_container_name";

            //Name of the Word file you want to load
            string blobName = "WordTemplate.docx";

            try
            {
                //Retrieve the document from Azure
                MemoryStream stream = await GetDocumentFromAzure(connectionString, containerName, blobName);

                //Set the position to the beginning of the MemoryStream
                stream.Position = 0;

                //Create an instance of WordDocument
                using (WordDocument wordDocument = new WordDocument(stream, Syncfusion.DocIO.FormatType.Docx))
                {
                    //Access the section in a Word document
                    IWSection section = wordDocument.Sections[0];

                    //Add new paragraph to the section
                    IWParagraph paragraph = section.AddParagraph();
                    paragraph.ParagraphFormat.FirstLineIndent = 36;
                    paragraph.BreakCharacterFormat.FontSize = 12f;

                    //Add new text to the paragraph
                    IWTextRange textRange = paragraph.AppendText("In 2000, AdventureWorks Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.") as IWTextRange;
                    textRange.CharacterFormat.FontSize = 12f;

                    //Saving the Word document to a MemoryStream
                    MemoryStream outputStream = new MemoryStream();
                    wordDocument.Save(outputStream, Syncfusion.DocIO.FormatType.Docx);

                    //Download the Word file in the browser
                    FileStreamResult fileStreamResult = new FileStreamResult(outputStream, "application/msword");
                    fileStreamResult.FileDownloadName = "EditWord.docx";
                    return fileStreamResult;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return Content("Error occurred while processing the file.");
            }
        }
        /// <summary>
        /// Download file from Azure Blob cloud storage
        /// </summary>
        /// <param name="bucketName"></param>
        /// <param name="key"></param>
        /// <returns></returns>
        public async Task<MemoryStream> GetDocumentFromAzure(string connectionString, string containerName, string blobName)
        {
            try
            {
                //Download the Word document from Azure Blob Storage
                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blobClient = containerClient.GetBlobClient(blobName);
                BlobDownloadInfo download = await blobClient.DownloadAsync();

                //Create a MemoryStream to copy the file content
                MemoryStream stream = new MemoryStream();
                await download.Content.CopyToAsync(stream);

                return stream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from Azure Blob Storage: {ex.Message}");
                throw; // or handle the exception as needed
            }
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
