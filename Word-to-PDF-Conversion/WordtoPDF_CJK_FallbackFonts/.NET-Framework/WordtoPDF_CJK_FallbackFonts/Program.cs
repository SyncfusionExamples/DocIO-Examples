using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Office;
using Syncfusion.Pdf;
namespace WordtoPDF_CJK_FallbackFonts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument(@"../../Data/Template.docx"))
            {
                document.FontSettings.FallbackFonts.Add(ScriptType.Chinese, "MicroSoft YaHei, SimSun");
                document.FontSettings.FallbackFonts.Add(ScriptType.Japanese, "Yu Gothic, Meiryo");
                document.FontSettings.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic");
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    PdfDocument pdf = converter.ConvertToPDF(document);
                    pdf.Save(@"../../Output/output.pdf");
                    pdf.Close(true);
                }
            }
        }
    }
}
