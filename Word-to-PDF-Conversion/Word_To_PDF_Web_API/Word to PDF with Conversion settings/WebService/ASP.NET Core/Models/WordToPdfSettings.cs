using Syncfusion.Pdf;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Runtime.InteropServices;

namespace WordToPdf.Models
{
    /// <summary>
    /// Provides settings for customizing Word to PDF conversion.
    /// </summary>
    public class WordToPdfSettings
    {
        /// <summary>
        /// The input Word document.
        /// </summary>
        [Required]
        public IFormFile? InputFile { get; set; }
        /// <summary>
        /// Specifies the password to open the protected input Word document.
        /// </summary>
        public string? Password { get; set; }
        /// <summary>
        /// Flag to embed the complete font information in the converted PDF document. The default value is false.
        /// </summary>
        [DefaultValue("false")]
        public bool? EmbedFontsInPDF { get; set; }
        /// <summary>
        /// Flag to preserve Word form fields as editable PDF form fields.
        /// </summary>
        [DefaultValue("true")]
        public bool? EditablePDF { get; set; } = true;
        /// <summary>
        /// Flag to automatically detect complex script text present in the Word document during PDF conversion. The default value is false.
        /// </summary>
        [DefaultValue("false")]
        public bool? AutoDetectComplexScript { get; set; } = false;
        /// <summary>
        /// Flag to convert the PDF document as tagged PDF or not. The default value is false.
        /// </summary>
        [DefaultValue("false")]
        public bool? TaggedPDF { get; set; } = false;

        /// <summary>
        /// Specifies the PDF document's conformance level.
        /// </summary>
        public PDFConformance? PdfConformanceLevel { get; set; } = PDFConformance.None;
        /// <summary>
        /// Flag to preserve the headings of the Word document as PDF bookmarks during the conversion process. The default value is true.
        /// </summary>
        [DefaultValue("true")]
        public bool? HeadingsAsPdfBookmarks { get; set; } = false;
        /// <summary>
        /// Flag to include comments from the Word document in the PDF.
        /// </summary>
        [DefaultValue("false")]
        public bool? IncludeComments { get; set; } = false;
        /// <summary>
        /// Flag to include revision of tracked changes Word document in the PDF.
        /// </summary>
        [DefaultValue("false")]
        public bool? IncludeRevisionsMarks { get; set; } = false;
    }
    /// <summary>
    /// Specifies the PDF document's conformance level.
    /// </summary>
    public enum PDFConformance
    {
        ///<summary>
        ///Specifies Default / No Conformance.
        ///</summary>
        None,
        ///<summary>
        ///This PDF/A ISO standard [ISO 19005-1:2005] is based on Adobe PDF version 1.4 and this Level B conformance indicates minimal compliance to ensure that the rendered visual appearance of a conforming file is preservable over the long term.
        ///</summary>
        Pdf_A1B,
        /// <summary>
        /// PDF/A-2 Standard is based on a PDF 1.7 (ISO 32000-1) which provides support for transparency effects and layers embedding of OpenType fonts
        /// </summary>
        Pdf_A2B,
        /// <summary>
        /// PDF/A-3 Standard is based on a PDF 1.7 (ISO 32000-1) which provides support for embedding the arbitrary file formats (XML, CSV, CAD, Word Processing documents)
        /// </summary>
        Pdf_A3B,
        /// <summary>
        /// This PDF/A ISO standard [ISO 19005-1:2005] is based on Adobe PDF version 1.4 and this Level A conformance was intended to increase the accessibility of conforming files for physically impaired users by allowing assistive software, such as screen readers, to more precisely extract and interpret a file's contents.
        /// </summary>
        Pdf_A1A,
        /// <summary>
        /// PDF/A-2 Standard is based on a PDF 1.7 (ISO 32000-1) and this Level A conformance was intended to increase the accessibility of conforming files for physically impaired users by allowing assistive software, such as screen readers, to more precisely extract and interpret a file's contents.
        /// </summary>
        Pdf_A2A,
        /// <summary>
        /// PDF/A-2 Standard is based on a PDF 1.7 (ISO 32000-1) and this Level U conformance represents Level B conformance (PDF/A-2b) with the additional requirement that all text in the document have Unicode mapping.
        /// </summary>
        Pdf_A2U,
        /// <summary>
        ///PDF/A-3 Standard is based on a PDF 1.7 (ISO 32000-1) which provides support for embedding the arbitrary file formats (XML, CSV, CAD, Word Processing documents) and This Level A conformance was intended to increase the accessibility of conforming files for physically impaired users by allowing assistive software, such as screen readers, to more precisely extract and interpret a file's contents.
        /// </summary>
        Pdf_A3A,
        /// <summary>
        /// PDF/A-3 Standard is based on a PDF 1.7 (ISO 32000-1) and this Level U conformance represents Level B conformance (PDF/A-3b) with the additional requirement that all text in the document have Unicode mapping.
        /// </summary>
        Pdf_A3U,
        /// <summary>
        /// PDF/A-4 Standard is based on a PDF 2.0 (ISO 32000-2). The separate conformance levels a, b, and u are not used in PDF/A-4. Instead, PDF/A-4 encourages but does not require the addition of higher-level logical structures, and it requires Unicode mappings for all fonts.
        /// </summary>
        Pdf_A4,
        /// <summary>
        /// PDF/A-4E Standard is based on a PDF 2.0 (ISO 32000-2). PDF/A-4e is intended for engineering documents and acts as a successor to the PDF/E-1 standard. PDF/A-4e supports Rich Media and 3D Annotations as well as embedded files.
        /// </summary>
        Pdf_A4E,
        /// <summary>
        /// PDF/A-4F Standard is based on a PDF 2.0 (ISO 32000-2). It allows embedding files in any other format.
        /// </summary>
        Pdf_A4F
    }
}
