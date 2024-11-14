using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Drawing;

using Amazon.Lambda.Core;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.Json.JsonSerializer))]

namespace MyLamdaProject
{
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
            //Path to the original library file.
            string originalLibraryPath = "/lib64/libdl.so.2";

            //Path to the symbolic link where the library will be copied.
            string symlinkLibraryPath = "/tmp/libdl.so";

            //Check if the original library file exists.
            if (File.Exists(originalLibraryPath))
            {
                //Copy the original library file to the symbolic link path, overwriting if it already exists.
                File.Copy(originalLibraryPath, symlinkLibraryPath, true);
            }

            string filePath = Path.GetFullPath(@"Data/Adventure.docx");

            //Load the file from the disk
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

            WordDocument document = new WordDocument(fileStream, FormatType.Docx);

            //Hooks the font substitution event
            document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;

            DocIORenderer render = new DocIORenderer();

            PdfDocument pdf = render.ConvertToPDF(document);

            //Unhooks the font substitution event after converting to PDF
            document.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;

            //Save the document into stream
            MemoryStream stream = new MemoryStream();

            //Save the PDF document  
            pdf.Save(stream);
            //Releases all resources used by the Word document and DocIO Renderer objects
            document.Close();
            render.Dispose();
            //Closes the PDF document
            pdf.Close();
            return Convert.ToBase64String(stream.ToArray());
        }
        /// <summary>
        /// Sets the alternate font when a specified font is not installed in the production environment
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/calibri.ttf"), FileMode.Open, FileAccess.Read);
            else if (args.OriginalFontName == "Arial" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/arial.ttf"), FileMode.Open, FileAccess.Read);
            else
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/times.ttf"), FileMode.Open, FileAccess.Read);
        }
    }
}
