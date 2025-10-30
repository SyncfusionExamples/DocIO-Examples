using System.Diagnostics;
using Create_IF_Field.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Create_IF_Field.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult SimpleIfField()
        {
            // Create a new Word document
            WordDocument document = new WordDocument();

            // Add a section and a title
            document.EnsureMinimal();
            IWParagraph titleParagraph = document.LastSection.AddParagraph();
            titleParagraph.AppendText("Student Grade Evaluation Report");
            titleParagraph.ApplyStyle(BuiltinStyle.Heading1);
            titleParagraph.ParagraphFormat.LineSpacing = 18f;

            // Add introductory explanation
            IWParagraph introParagraph = document.LastSection.AddParagraph();
            introParagraph.AppendText("This report evaluates a student's performance based on their exam score. "
                + "If the score meets or exceeds the passing mark, the student is considered to have passed. "
                + "Otherwise, the student is marked as failed.");
            introParagraph.ParagraphFormat.LineSpacing = 18f;

            // Add details of the evaluation criteria
            IWParagraph criteriaParagraph = document.LastSection.AddParagraph();
            criteriaParagraph.AppendText("Evaluation Criteria:\n"
                + "• Passing Mark: 50\n"
                + "• Student Score: 75\n"
                + "• Condition: IF Student Score >= Passing Mark");
            criteriaParagraph.ParagraphFormat.LineSpacing = 18f;

            // Add a paragraph with the IF field
            IWParagraph resultParagraph = document.LastSection.AddParagraph();
            resultParagraph.AppendText("Evaluation Result: ");
            resultParagraph.ParagraphFormat.LineSpacing = 18f;

            // Add the IF field to show pass/fail result
            WIfField? gradeIfField = resultParagraph.AppendField("If", FieldType.FieldIf) as WIfField;
            gradeIfField!.FieldCode = "IF 75 >= 50 \"Pass ✅\" \"Fail ❌\"";

            // Update fields to evaluate the IF condition
            document.UpdateDocumentFields();

            // Return the document as a downloadable file
            return CreateFileResult(document, "SimpleIfField.docx");
        }

        public IActionResult IfFieldWithRichContent()
        {
            // Load an existing Word document from the specified file path
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Sample.docx"));

            // Get the last paragraph in the document to append the IF field
            WParagraph paragraph = document.LastParagraph;

            // Define the variable value to be used in the IF condition
            string product = "cycle";

            // Append an IF field to the paragraph
            WField field = (WField)paragraph.AppendField("If", FieldType.FieldIf);

            // Insert field code and rich content (tables and text) for true/false statements
            InsertIfFieldCode(paragraph, field, product);

            // Update all fields in the document
            document.UpdateDocumentFields();

            // Return the generated Word document as a downloadable file
            return CreateFileResult(document, "IfFieldWithRichContent.docx");
        }

        /// <summary>
        /// Insert IF field code with complex true/false content (tables and text)
        /// </summary>
        private static void InsertIfFieldCode(WParagraph paragraph, WField field, string product)
        {
            // Get the index of the field in the paragraph
            int fieldIndex = paragraph.Items.IndexOf(field) + 1;

            // Define the IF field condition
            field.FieldCode = $"IF \"{product}\" = \"cycle\" ";

            // Move field separator and end marks to a temporary paragraph
            WParagraph lastPara = new WParagraph(paragraph.Document);
            MoveFieldMark(paragraph, fieldIndex + 1, lastPara);

            // Insert true statement (when condition is true)
            paragraph = InsertTrueStatement(paragraph);

            // Insert false statement (when condition is false)
            paragraph = InsertFalseStatement(paragraph);

            // Move remaining field marks back from temporary paragraph to the original
            MoveFieldMark(lastPara, 0, paragraph);
        }

        /// <summary>
        /// Moves remaining field items to another paragraph
        /// </summary>
        private static void MoveFieldMark(WParagraph paragraph, int fieldIndex, WParagraph lastPara)
        {
            // Move all items after the field index to the destination paragraph
            for (int i = fieldIndex; i < paragraph.Items.Count;)
                lastPara.Items.Add(paragraph.Items[i]);
        }

        /// <summary>
        /// Insert the true part of the IF field with rich content (text + table)
        /// </summary>
        private static WParagraph InsertTrueStatement(WParagraph paragraph)
        {
            WTextBody ownerTextBody = paragraph.OwnerTextBody;

            // Append heading text for the true statement
            WTextRange text = (WTextRange)paragraph.AppendText("\"Product Overview");
            text.CharacterFormat.Bold = true;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a table with product details
            WTable table = (WTable)ownerTextBody.AddTable();

            // Add rows and cells to the table with data
            WTableRow row = table!.AddRow() as WTableRow;
            row.AddCell().AddParagraph().AppendText("Mountain-200");
            row.AddCell().AddParagraph().AppendText("$2,294.99");

            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("Mountain-300");
            row.Cells[1].AddParagraph().AppendText("$1,079.99");

            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("Road-150");
            row.Cells[1].AddParagraph().AppendText("$3,578.27");

            // Add a paragraph to close the true statement string
            WParagraph lastPara = (WParagraph)ownerTextBody.AddParagraph();
            lastPara.AppendText("\" ");
            return lastPara;
        }

        /// <summary>
        /// Insert the false part of the IF field with rich content (text + table)
        /// </summary>
        private static WParagraph InsertFalseStatement(WParagraph paragraph)
        {
            WTextBody ownerTextBody = paragraph.OwnerTextBody;

            // Append heading text for the false statement
            WTextRange text = (WTextRange)paragraph.AppendText("\"Juice Corner");
            text.CharacterFormat.Bold = true;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a table with juice product details
            WTable table = (WTable)ownerTextBody.AddTable();

            // Add rows and cells to the table with data
            WTableRow row = table.AddRow() as WTableRow;
            row.AddCell().AddParagraph().AppendText("Apple Juice");
            row.AddCell().AddParagraph().AppendText("$12.00");

            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("Grape Juice");
            row.Cells[1].AddParagraph().AppendText("$15.00");

            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("Hot Soup");
            row.Cells[1].AddParagraph().AppendText("$20.00");

            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("Tender Coconut");
            row.Cells[1].AddParagraph().AppendText("$20.00");

            row = table.AddRow() as WTableRow;
            row.Cells[0].AddParagraph().AppendText("Cherry");
            row.Cells[1].AddParagraph().AppendText("$25.00");

            // Add a paragraph to close the false statement string
            WParagraph lastPara = (WParagraph)ownerTextBody.AddParagraph();
            lastPara.AppendText("\" ");
            return lastPara;
        }

        private FileStreamResult CreateFileResult(WordDocument document, string fileName)
        {
            MemoryStream outputStream = new MemoryStream();
            document.Save(outputStream, FormatType.Docx);
            outputStream.Position = 0;
            return File(outputStream, "application/docx", fileName);
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
