using Google.Apis.Auth.OAuth2;
using Google.Cloud.Storage.V1;
using Microsoft.AspNetCore.Mvc;
using Open_Word_document.Models;
using Syncfusion.DocIO.DLS;
using System.Diagnostics;

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
            //Download the file from Google
            MemoryStream memoryStream = await GetDocumentFromGoogle();

            //Create an instance of WordDocument
            using (WordDocument wordDocument = new WordDocument(memoryStream, Syncfusion.DocIO.FormatType.Docx))
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
        /// <summary>
        /// Download file from Google
        /// </summary>
        /// <returns></returns>
        public async Task<MemoryStream> GetDocumentFromGoogle()
        {
            try
            {
                //Your bucket name
                string bucketName = "Your_bucket_name";

                //Your service account key path
                string keyPath = "Your_service_account_key_path";

                //Name of the file to download from the Google Cloud Storage
                string fileName = "WordTemplate.docx";

                //Create Google Credential from the service account key file
                GoogleCredential credential = GoogleCredential.FromFile(keyPath);

                //Instantiates a storage client to interact with Google Cloud Storage
                StorageClient storageClient = StorageClient.Create(credential);

                //Download a file from Google Cloud Storage
                MemoryStream memoryStream = new MemoryStream();
                await storageClient.DownloadObjectAsync(bucketName, fileName, memoryStream);
                memoryStream.Position = 0;

                return memoryStream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from Google Cloud Storage: {ex.Message}");
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
