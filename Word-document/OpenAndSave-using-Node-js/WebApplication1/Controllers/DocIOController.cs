using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Cors;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace WebApplication1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocIOController : Controller
    {
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("OpenAndResave")]
        public IActionResult OpenAndResave(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Invalid file uploaded");

            try
            {
                using (Stream inputStream = file.OpenReadStream())
                {
                    // Open the Word document
                    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                    {
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            // Save the document to the memory stream
                            document.Save(memoryStream, FormatType.Docx);

                            // Reset the stream position to the beginning
                            memoryStream.Position = 0;
                            // Convert MemoryStream to Byte Array
                            byte[] fileBytes = memoryStream.ToArray();
                            // Return the file as a downloadable response
                            return File(memoryStream.ToArray(),
                                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        "ResavedDocument.docx");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error processing document: {ex.Message}");
            }
        }
    }
}
