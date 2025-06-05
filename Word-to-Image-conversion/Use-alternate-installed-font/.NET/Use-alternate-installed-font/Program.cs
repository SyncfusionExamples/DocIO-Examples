using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Use_alternate_installed_font
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Hooks the font substitution event.
                    wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert the entire Word document to images.
                        Stream[] imageStreams = wordDocument.RenderAsImages();
                        int i = 0;
                        foreach (Stream stream in imageStreams)
                        {
                            //Reset the stream position.
                            stream.Position = 0;
                            //Save the stream as file.
                            using (FileStream fileStreamOutput = File.Create("Output/WordToImage_" + i + ".jpeg"))
                            {
                                stream.CopyTo(fileStreamOutput);
                            }
                            i++;
                        }
                    }
                    //Unhooks the font substitution event after converting to Image.
                    wordDocument.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                }
            }
        }
        private static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            //Sets the alternate font when a specified font is not installed in the production environment.
            //If "Arial Unicode MS" font is not installed, then it uses the "Arial" font.
            //For other missing fonts, uses the "Times New Roman".
            if (args.OriginalFontName == "Arial Unicode MS")
                args.AlternateFontName = "Arial";
            else
                args.AlternateFontName = "Times New Roman";
        }
    }
}
