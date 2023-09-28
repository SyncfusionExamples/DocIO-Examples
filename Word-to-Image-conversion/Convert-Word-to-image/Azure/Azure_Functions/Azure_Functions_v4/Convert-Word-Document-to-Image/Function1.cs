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
using Syncfusion.DocIORenderer;
using Syncfusion.DocIO;
using System.Net.Http;
using System.Net;
using Microsoft.Azure.WebJobs.Host;
using System.Net.Http.Headers;

namespace Convert_Word_Document_to_Image
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input Word document as stream from request
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Loads an existing Word document
            using (WordDocument document = new WordDocument(stream, FormatType.Docx))
            {
                //Creates an instance of the DocIORenderer
                using (DocIORenderer render = new DocIORenderer())
                {
                    Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    imageStream.Position = 0;
                    MemoryStream memoryStream = new MemoryStream();
                    //Saves the Image file 
                    imageStream.CopyTo(memoryStream);
                    //Reset the memory stream position
                    memoryStream.Position = 0;
                    //Create the response to return
                    HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                    //Set the Word document saved stream as content of response
                    response.Content = new ByteArrayContent(memoryStream.ToArray());
                    //Set the contentDisposition as attachment
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = "WordToImage.Jpeg"
                    };
                    //Set the content type as Word document mime type
                    response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/jpeg");
                    //Return the response with output Word document stream
                    return response;
                }
            }
        }
    }
}
