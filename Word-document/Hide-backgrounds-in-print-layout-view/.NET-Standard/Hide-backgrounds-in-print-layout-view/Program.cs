using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Hide_backgrounds_in_print_layout_view
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Sets the background type as gradient.
                    document.Background.Type = BackgroundType.Gradient;
                    //Set color for gradient.
                    document.Background.Gradient.Color1 = Syncfusion.Drawing.Color.LightGray;
                    document.Background.Gradient.Color2 = Syncfusion.Drawing.Color.LightGreen;
                    //Set the shading style.
                    document.Background.Gradient.ShadingStyle = GradientShadingStyle.DiagonalUp;
                    document.Background.Gradient.ShadingVariant = GradientShadingVariant.ShadingDown;
                    //Set whether background colors and images are shown when a document is displayed in print layout view.
                    document.Settings.DisplayBackgrounds = false;
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
