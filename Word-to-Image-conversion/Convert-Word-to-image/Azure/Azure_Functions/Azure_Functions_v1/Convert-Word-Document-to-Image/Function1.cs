using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Convert_Word_Document_to_Image
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input Word document as stream from request.
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Loads an existing Word document
            using (WordDocument document = new WordDocument(stream))
            {
                //Convert the first page of the Word document into an image.
                System.Drawing.Image image = document.RenderAsImages(0, ImageType.Bitmap);
                //initializes a new instance of the MemoryStream.
                MemoryStream memoryStream = new MemoryStream();
                //Saves the Image file.
                image.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                //Reset the memory stream position.
                memoryStream.Position = 0;
                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the Word document saved stream as content of response.
                response.Content = new ByteArrayContent(memoryStream.ToArray());
                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "wordtoimage.jpeg"
                };
                //Set the content type as Word document mime type.
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/jpeg");
                //Return the response with output Word document stream.
                return response;
            }
        }
    }
}
