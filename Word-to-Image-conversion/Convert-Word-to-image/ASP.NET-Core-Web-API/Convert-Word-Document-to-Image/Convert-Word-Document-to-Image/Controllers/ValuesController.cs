using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Convert_Word_Document_to_Image.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("api/ConvertWordToImage")]
        public IActionResult ConvertWordToImage()
        {
            try
            {
                var fileDownloadName = "Output.jpeg";
                const string contentType = "image/jpeg";
                var stream = ConvertWordDocumentToImage();
                stream.Position = 0;
                return File(stream, contentType, fileDownloadName);
            }
            catch (Exception ex)
            {
                return BadRequest("Error occurred while converting Word to Image: " + ex.Message);
            }
        }
        public static Stream ConvertWordDocumentToImage()
        {
            //Loads the input Word document
            WordDocument wordDocument = new WordDocument(Path.GetFullPath("Data/Input.docx"), FormatType.Docx);   
            DocIORenderer render = new DocIORenderer();
            //Convert the first page of the Word document into an image.
            Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
            //close the word document.
            wordDocument.Close();
            //Reset the stream position.
            imageStream.Position = 0;
            //Save the image file.
            return imageStream;                    
        }
    }
}
