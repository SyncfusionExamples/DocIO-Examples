using Dropbox.Api;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
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
            try
            {
                //Retrieve the document from DropBox
                MemoryStream stream = await GetDocumentFromDropBox();

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
        /// Download file from DropBox cloud storage
        /// </summary>
        /// <returns></returns>
        public async Task<MemoryStream> GetDocumentFromDropBox()
        {
            //Define the access token for authentication with the Dropbox API
            var accessToken = "Access_Token";

            //Define the file path in Dropbox where the file is located
            var filePathInDropbox = "FilePath";

            try
            {
                //Create a new DropboxClient instance using the provided access token
                using (var dbx = new DropboxClient(accessToken))
                {
                    //Start a download request for the specified file in Dropbox
                    using (var response = await dbx.Files.DownloadAsync(filePathInDropbox))
                    {
                        //Get the content of the downloaded file as a stream
                        var content = await response.GetContentAsStreamAsync();

                        MemoryStream stream = new MemoryStream();
                        content.CopyTo(stream);
                        return stream;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from DropBox: {ex.Message}");
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
