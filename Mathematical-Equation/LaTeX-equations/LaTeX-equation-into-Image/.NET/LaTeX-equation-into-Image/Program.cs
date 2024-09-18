using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

// Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add one section and one paragraph to the document.
    document.EnsureMinimal();

    //Append an accent equation using LaTeX.
    document.LastParagraph.AppendMath(@"f\left(x\right)={a}_{0}+\sum_{n=1}^{\infty}{\left({a}_{n}\cos{\frac{n\pi{x}}{L}}+{b}_{n}\sin{\frac{n\pi{x}}{L}}\right)}");
    
    //Instantiation of DocIORenderer for Word to PDF conversion
    DocIORenderer render = new DocIORenderer();
    //Convert the first page of the Word document into an image.
    Stream imageStream = document.RenderAsImages(0, ExportImageFormat.Jpeg);
    //Reset the stream position.
    imageStream.Position = 0;
    //Save the stream as file.
    using (FileStream fileStreamOutput = File.Create("Output/Result.jpeg"))
    {
        imageStream.CopyTo(fileStreamOutput);
    }
    //Close the Word document
    document.Close();
}