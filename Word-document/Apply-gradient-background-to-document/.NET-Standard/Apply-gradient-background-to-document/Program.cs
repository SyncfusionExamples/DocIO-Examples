using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Apply_gradient_background_to_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an document from file system through constructor of WordDocument class.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Sets the background type as gradient.
                    document.Background.Type = BackgroundType.Gradient;
                    //Sets color for gradient.
                    document.Background.Gradient.Color1 = Syncfusion.Drawing.Color.LightGray;
                    document.Background.Gradient.Color2 = Syncfusion.Drawing.Color.LightGreen;
                    //Sets the shading style.
                    document.Background.Gradient.ShadingStyle = GradientShadingStyle.DiagonalUp;
                    document.Background.Gradient.ShadingVariant = GradientShadingVariant.ShadingDown;
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
