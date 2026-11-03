using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace WordtoPDF.Services;

public class PdfConverter
{
    public MemoryStream Convert(Stream input, string fileName)
    {
        var ext = Path.GetExtension(fileName);

        if (string.Equals(ext, ".docx", StringComparison.OrdinalIgnoreCase))
        {
            using var document = new WordDocument(input, Syncfusion.DocIO.FormatType.Docx);
            using var renderer = new DocIORenderer();
            using PdfDocument pdf = renderer.ConvertToPDF(document);

            var output = new MemoryStream();
            pdf.Save(output);
            output.Position = 0;
            return output;
        }
        else if (string.Equals(ext, ".pptx", StringComparison.OrdinalIgnoreCase))
        {
            using var presentation = Presentation.Open(input);
            using PdfDocument pdf = PresentationToPdfConverter.Convert(presentation);

            var output = new MemoryStream();
            pdf.Save(output);
            output.Position = 0;
            return output;
        }
        else
        {
            throw new NotSupportedException($"File extension '{ext}' is not supported. Only .docx and .pptx are allowed.");
        }
    }
}
