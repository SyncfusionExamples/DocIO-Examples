using Syncfusion.DocIO.DLS;

namespace Syncfusion.Formatting;

/// <summary>
/// Defines the formatting standards used to normalize document content.
/// Any content that does not match these rules is considered inconsistent
/// and can be updated during document processing.
/// </summary>
public sealed class FormattingRules
{
    /// <summary>
    /// Default font family applied to regular document text.
    /// </summary>
    public string DefaultFontFamily { get; init; } = "Calibri";

    /// <summary>
    /// Default font family applied to Arabic text.
    /// </summary>
    public string DefaultArabicFontFamily { get; init; } = "Arial";

    /// <summary>
    /// Default font size, in points, applied to document text.
    /// </summary>
    public float DefaultFontSize { get; init; } = 11f;

    /// <summary>
    /// Font families that are exempt from font normalization.
    /// These fonts are preserved when encountered in the document.
    /// Comparison is case-insensitive.
    /// </summary>
    public IReadOnlyCollection<string> AllowedFontFamilies { get; init; } = Array.Empty<string>();

    /// <summary>
    /// Default spacing before a paragraph, in points.
    /// </summary>
    public float DefaultParagraphBeforeSpacing { get; init; } = 0f;

    /// <summary>
    /// Default spacing after a paragraph, in points.
    /// </summary>
    public float DefaultParagraphAfterSpacing { get; init; } = 8f;

    /// <summary>
    /// Default horizontal alignment for paragraphs.
    /// </summary>
    public HorizontalAlignment DefaultParaAlignment { get; init; } = HorizontalAlignment.Left;

    // ---------------------------------------------------------------------
    // Table Cell Formatting
    // ---------------------------------------------------------------------

    /// <summary>
    /// Default top cell padding, in points.
    /// </summary>
    public float DefaultCellPaddingTop { get; init; } = 0f;

    /// <summary>
    /// Default bottom cell padding, in points.
    /// </summary>
    public float DefaultCellPaddingBottom { get; init; } = 0f;

    /// <summary>
    /// Default left cell padding, in points.
    /// </summary>
    public float DefaultCellPaddingLeft { get; init; } = 5.4f;

    /// <summary>
    /// Default right cell padding, in points.
    /// </summary>
    public float DefaultCellPaddingRight { get; init; } = 5.4f;

    /// <summary>
    /// Default horizontal alignment for text within table cells.
    /// </summary>
    public HorizontalAlignment DefaultCellTextAlignment { get; init; } = HorizontalAlignment.Left;

    /// <summary>
    /// Default spacing before paragraphs contained in table cells, in points.
    /// </summary>
    public float DefaultCellParagraphSpacingBefore { get; init; } = 0f;

    /// <summary>
    /// Default spacing after paragraphs contained in table cells, in points.
    /// </summary>
    public float DefaultCellParagraphSpacingAfter { get; init; } = 0f;

    // ---------------------------------------------------------------------
    // Table Formatting
    // ---------------------------------------------------------------------

    /// <summary>
    /// Default left indentation applied to tables, in points.
    /// </summary>
    public float DefaultTableLeftIndent { get; init; } = 0f;

    /// <summary>
    /// Indicates whether text wrapping is enabled around tables.
    /// </summary>
    public bool DefaultTableTextWrap { get; init; } = false;

    /// <summary>
    /// Default top padding for tables, in points.
    /// </summary>
    public float DefaultTablePaddingTop { get; init; } = 0f;

    /// <summary>
    /// Default bottom padding for tables, in points.
    /// </summary>
    public float DefaultTablePaddingBottom { get; init; } = 0f;

    /// <summary>
    /// Default left padding for tables, in points.
    /// </summary>
    public float DefaultTablePaddingLeft { get; init; } = 5.4f;

    /// <summary>
    /// Default right padding for tables, in points.
    /// </summary>
    public float DefaultTablePaddingRight { get; init; } = 5.4f;

    // ---------------------------------------------------------------------
    // Image Formatting
    // ---------------------------------------------------------------------

    /// <summary>
    /// Default image height, in points.
    /// </summary>
    public float DefaultImageHeight { get; init; } = 150f;

    /// <summary>
    /// Default image width, in points.
    /// </summary>
    public float DefaultImageWidth { get; init; } = 150f;
}