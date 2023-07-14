using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using Microsoft.Azure.WebJobs.Host;

namespace Open_and_Save_Word_Document
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input Word document as stream from request
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Loads an existing Word document
            using (WordDocument document = new WordDocument(stream,FormatType.Docx))
            {
                //Access the section in a Word document.
                IWSection section = document.Sections[0];
                //Add a new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                IWTextRange text = paragraph.AppendText("In 2000, Adventure Works Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the Adventure Works Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.");
                text.CharacterFormat.FontSize = 12f;

                MemoryStream memoryStream = new MemoryStream();
                //Saves the Word document file.
                document.Save(memoryStream, FormatType.Docx);
                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the Word document saved stream as content of response.
                response.Content = new ByteArrayContent(memoryStream.ToArray());
                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Sample.docx"
                };
                //Set the content type as Word document mime type.
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/docx");
                //Return the response with output Word document stream.
                return response;
            }
        }
    }
}
