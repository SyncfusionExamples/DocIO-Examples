using Amazon.Lambda.Core;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;

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
                //Hooks the font substitution event.
                wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                //Instantiation of DocIORenderer.
                using (DocIORenderer render = new DocIORenderer())
                {
                    //Convert the first page of the Word document into an image.
                    Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
                    //Unhooks the font substitution event after converting to image.
                    wordDocument.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
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

    //Set the alternate font when a specified font is not installed in the production environment.
    private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
    {
        if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
            args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/calibri.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        else if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Bold)
            args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/calibrib.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        else
            args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/times.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }
}
