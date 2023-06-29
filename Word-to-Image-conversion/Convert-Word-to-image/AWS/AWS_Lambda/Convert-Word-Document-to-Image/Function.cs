using Amazon.Lambda.Core;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Convert_Word_Document_to_Image;

public class Function
{
    
    /// <summary>
    /// A simple function that takes a string and does a ToUpper
    /// </summary>
    /// <param name="input"></param>
    /// <param name="context"></param>
    /// <returns></returns>
    public string FunctionHandler(string input, ILambdaContext context)
    {
        string filePath = Path.GetFullPath(@"Data/Input.docx");
        //Open the file as Stream.
        using (FileStream docStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            //Loads file stream into Word document.
            using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
            {
                //Instantiation of DocIORenderer.
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Convert the first page of the Word document into an image.
                    Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    imageStream.Position = 0;
                    //Save the image file into stream.
                    MemoryStream stream = new MemoryStream();
                    imageStream.CopyTo(stream);
                    return Convert.ToBase64String(stream.ToArray());
                }
            }
        }
    }
}
