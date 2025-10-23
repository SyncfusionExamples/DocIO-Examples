using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Reflection.Metadata;

namespace Open_and_save_Word_document.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("api/Word")]
        public IActionResult DownloadWordDocument()
        {
            try
            {
                var fileDownloadName = "Output.docx";
                const string contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                var stream = OpenandSaveDocument();
                stream.Position = 0;
                return File(stream, contentType, fileDownloadName);
            }
            catch (Exception ex)
            {
                // Log or handle the exception
                return BadRequest("Error occurred while creating Word file: " + ex.Message);
            }
        }
        public static MemoryStream OpenandSaveDocument()
        {
            //Open an existing Word document.
            WordDocument document = new WordDocument(Path.GetFullPath("Data/Input.docx"));
            //Access the section in a Word document.
            IWSection section = document.Sections[0];
            //Add a new paragraph to the section.
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.FirstLineIndent = 36;
            paragraph.BreakCharacterFormat.FontSize = 12f;
            IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
            text.CharacterFormat.FontSize = 12f;

            //Saving the Word document to the MemoryStream 
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            document.Close();
            //Set the position as '0'.
            stream.Position = 0;
            return stream;
        }
    }
}
