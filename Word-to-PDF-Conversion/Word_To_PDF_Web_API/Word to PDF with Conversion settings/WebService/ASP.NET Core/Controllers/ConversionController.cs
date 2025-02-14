using WordToPdf.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;
using System.Reflection.Metadata;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Cors;
using Microsoft.IdentityModel.Tokens;

namespace WordToPdf.Controllers
{
    /// <summary>
    /// Controller for Word to PDF conversion.
    /// </summary>
    [Route("api/pdf/")]
    [ApiController]
    [EnableCors("MyPolicy")]
    public class ConversionController : ControllerBase
    {
        #region Fields
        /// <summary>
        /// Hosting environment information.
        /// </summary>
        private readonly IWebHostEnvironment _webHostEnvironment;
        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the ConversionController class with the specified hosting environment.
        /// </summary>
        /// <param name="webHostEnvironment">Represents the hosting environment of the application.</param>
        public ConversionController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        #endregion

        #region Web API
        /// <summary>
        /// Convert Word document to PDF.
        /// </summary>
        /// <param name="settings">Settings to customize the Word to PDF conversion.</param>
        [HttpPost]
        [Route("convertwordtopdf")]
        public IActionResult WordToPDFConversion([FromForm] WordToPdfSettings settings)
        {
            //Convert Word document to PDF
            return WordToPdf(settings);
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// Convert Word document to PDF with the customization settings.
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        private IActionResult WordToPdf(WordToPdfSettings settings)
        {
            try
            {
                if (settings.InputFile != null)
                {

                    //Get input file
                    MemoryStream docStream = new MemoryStream();
                    settings.InputFile.CopyTo(docStream);
                    docStream.Position = 0;

                    WordDocument wordDocument;
                    //Load Word document.
                    if (string.IsNullOrEmpty(settings.Password))
                    {
                        wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic);
                    }
                    else
                    {
                        wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic, settings.Password);
                    }
                    docStream.Dispose();

                    //Instantiation of DocIORenderer for Word to PDF conversion.
                    DocIORenderer render = new DocIORenderer();

                    //Apply settings for Word to PDF conversion.
                    ApplyWordToPDFSettings(render.Settings, settings, wordDocument);

                    //Converts Word document into PDF document
                    PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);

                    //Dispose the resources.
                    render.Dispose();
                    wordDocument.Dispose();

                    //Saves the PDF document to MemoryStream.
                    MemoryStream stream = new MemoryStream();
                    pdfDocument.Save(stream);
                    stream.Position = 0;
                    //Dispose the PDF resources.
                    pdfDocument.Close(true);
                    PdfDocument.ClearFontCache();

                    //Return the PDF document.
                    return File(stream, "application/pdf", "OutputFile.pdf");
                }

            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred during the conversion process: {ex.Message}");
            }

            return StatusCode(500, "An error occurred during Word to PDF conversion.");
        }

        /// <summary>
        /// Applies the settings for Word to PDF conversion from the Word to PDF API options.
        /// </summary>
        /// <param name="rendererSettings">The settings for the DocIORenderer.</param>
        /// <param name="settings">The Word to PDF API options.</param>
        /// <param name="document">The Word document to be converted.</param>
        internal static void ApplyWordToPDFSettings(DocIORendererSettings rendererSettings, WordToPdfSettings settings, WordDocument document)
        {
            //Set the flag to optimize identical images during Word to PDF conversion.
            rendererSettings.OptimizeIdenticalImages = true;

            //Set the flag to generate tagged PDF from the Word document.
            if (settings.TaggedPDF.HasValue)
                rendererSettings.AutoTag = settings.TaggedPDF.Value;

            //Set the revision markups to include the revision of tracked changes in the Word document in the converted PDF.
            if (settings.IncludeRevisionsMarks.HasValue && settings.IncludeRevisionsMarks.Value)
                document.RevisionOptions.ShowMarkup = RevisionType.Deletions | RevisionType.Formatting | RevisionType.Insertions;

            //Set the comment display mode as balloons to render Word document comments in the converted PDF.
            if (settings.IncludeComments.HasValue && settings.IncludeComments.Value)
                document.RevisionOptions.CommentDisplayMode = CommentDisplayMode.ShowInBalloons;

            //Set the flag to preserve text form fields as editable PDF form fields.
            if (settings.EditablePDF.HasValue)
                rendererSettings.PreserveFormFields = settings.EditablePDF.Value;

            //Set the flag to embed fonts in the converted PDF.
            if (settings.EmbedFontsInPDF.HasValue && settings.EmbedFontsInPDF.Value)
                rendererSettings.EmbedCompleteFonts = settings.EmbedFontsInPDF.Value;

            //Set the flag to auto-detect complex script during Word to PDF conversion.
            if (settings.AutoDetectComplexScript.HasValue && settings.AutoDetectComplexScript.Value)
                rendererSettings.AutoDetectComplexScript = settings.AutoDetectComplexScript.Value;

            //Set the PDF conformance level.
            rendererSettings.PdfConformanceLevel = GetPdfConformanceLevel(settings.PdfConformanceLevel);

            //Set the option to generate a PDF document with bookmarks for Word document paragraphs with a heading style and outline level.
            if (settings.HeadingsAsPdfBookmarks.HasValue)
                rendererSettings.ExportBookmarks = settings.HeadingsAsPdfBookmarks.Value ? Syncfusion.DocIO.ExportBookmarkType.Headings | Syncfusion.DocIO.ExportBookmarkType.Bookmarks:
                    Syncfusion.DocIO.ExportBookmarkType.Bookmarks;
        }
        /// <summary>
        /// Convert the PDF conformance from Web API to its equivalent <see cref="Syncfusion.Pdf.PdfConformanceLevel"/>
        /// </summary>
        /// <param name="pdfConformance">PDF conformance from Web API.</param>
        /// <returns>Returns <see cref="Syncfusion.Pdf.PdfConformanceLevel"/></returns>
        private static PdfConformanceLevel GetPdfConformanceLevel(PDFConformance? pdfConformance)
        {
            if (!pdfConformance.HasValue)
                return PdfConformanceLevel.None;

            switch (pdfConformance)
            {
                case PDFConformance.None:
                    return PdfConformanceLevel.None;
                case PDFConformance.Pdf_A1B:
                    return PdfConformanceLevel.Pdf_A1B;
                case PDFConformance.Pdf_A2B:
                    return PdfConformanceLevel.Pdf_A2B;
                case PDFConformance.Pdf_A3B:
                    return PdfConformanceLevel.Pdf_A3B;
                case PDFConformance.Pdf_A1A:
                    return PdfConformanceLevel.Pdf_A1A;
                case PDFConformance.Pdf_A2A:
                    return PdfConformanceLevel.Pdf_A2A;
                case PDFConformance.Pdf_A2U:
                    return PdfConformanceLevel.Pdf_A2U;
                case PDFConformance.Pdf_A3A:
                    return PdfConformanceLevel.Pdf_A3A;
                case PDFConformance.Pdf_A3U:
                    return PdfConformanceLevel.Pdf_A3U;
                case PDFConformance.Pdf_A4:
                    return PdfConformanceLevel.Pdf_A4;
                case PDFConformance.Pdf_A4E:
                    return PdfConformanceLevel.Pdf_A4E;
                case PDFConformance.Pdf_A4F:
                    return PdfConformanceLevel.Pdf_A4F;
                default:
                    return PdfConformanceLevel.None;
            }
        }
        #endregion
    }
}
