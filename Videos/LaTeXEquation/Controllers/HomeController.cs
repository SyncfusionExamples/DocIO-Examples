using System.Diagnostics;
using LaTeXEquation.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace LaTeXEquation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        // Creates a new Word document with mathematical equations
        public IActionResult CreateEquation()
        {
            // Initialize a new Word document and ensure minimal structure
            WordDocument document = new WordDocument();
            document.EnsureMinimal();

            // Add a main title paragraph
            IWParagraph mainTitle = document.LastSection.AddParagraph();
            IWTextRange titleText = mainTitle.AppendText("Mathematical Equations");
            titleText.CharacterFormat.Bold = true; // Make title bold
            titleText.CharacterFormat.FontSize = 18; // Set font size
            mainTitle.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center; // Center align title
            mainTitle.ParagraphFormat.AfterSpacing = 12f; // Add spacing after title

            // Add paragraph for Area of Circle equation
            IWParagraph paragraph1 = document.LastSection.AddParagraph();
            paragraph1.ParagraphFormat.BeforeSpacing = 8f; // Add spacing before paragraph
            IWTextRange wText1 = paragraph1.AppendText("Area of Circle ");
            wText1.CharacterFormat.Bold = true; // Make label bold
            paragraph1.AppendMath(@"A = \pi r^2"); // Insert LaTeX math equation

            // Add paragraph for Quadratic Formula equation
            IWParagraph paragraph2 = document.LastSection.AddParagraph();
            paragraph2.ParagraphFormat.BeforeSpacing = 8f;
            IWTextRange wText2 = paragraph2.AppendText("Quadratic Formula ");
            wText2.CharacterFormat.Bold = true;
            paragraph2.AppendMath(@"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}");

            // Add paragraph for Fourier Series equation
            IWParagraph paragraph3 = document.LastSection.AddParagraph();
            paragraph3.ParagraphFormat.BeforeSpacing = 8f;
            IWTextRange wText3 = paragraph3.AppendText("Fourier Series ");
            wText3.CharacterFormat.Bold = true;
            paragraph3.AppendMath(@"f\left(x\right)={a}_{0}+\sum_{n=1}^{\infty}{\left({a}_{n}\cos{\frac{n\pi{x}}{L}}+{b}_{n}\sin{\frac{n\pi{x}}{L}}\right)}");

            // Return the document as a downloadable file
            return CreateFileResult(document, "CreateEquation.docx");
        }

        // Adds a new equation to an existing Word document
        public IActionResult AddEquation()
        {
            // Open existing document from file
            using (FileStream fileStream = new FileStream("Data\\Input.docx", FileMode.Open, FileAccess.Read))
            {
                WordDocument document = new WordDocument(fileStream, FormatType.Automatic);

                // Find the text "Derivative equation" in the document
                TextSelection selection = document.Find("Derivative equation", false, true);

                if (selection != null)
                {
                    // Get the paragraph containing the found text
                    WParagraph targetParagraph = selection.GetAsOneRange().OwnerParagraph as WParagraph;

                    if (targetParagraph != null)
                    {
                        // Create a new paragraph for the math equation
                        WParagraph newParagraph = new WParagraph(document);

                        // Create a math object and set its LaTeX representation
                        WMath math = new WMath(document);
                        math.MathParagraph.LaTeX = @"\frac{d}{dx}\left(x^n\right)=nx^{n-1}";

                        // Add the math object to the new paragraph
                        newParagraph.ChildEntities.Add(math);

                        // Insert the new paragraph after the target paragraph
                        int index = document.LastSection.Body.ChildEntities.IndexOf(targetParagraph);
                        document.LastSection.Body.ChildEntities.Insert(index + 1, newParagraph);
                    }
                }

                // Return the updated document as a downloadable file
                return CreateFileResult(document, "AddEquation.docx");
            }
        }

        // Edits an existing equation in a Word document
        public IActionResult EditEquation()
        {
            // Open template document from file
            using (FileStream fileStream = new FileStream("Data\\Template.docx", FileMode.Open, FileAccess.Read))
            {
                WordDocument document = new WordDocument(fileStream, FormatType.Automatic);

                // Find the first math object in the document
                WMath? math = document.FindItemByProperty(EntityType.Math, string.Empty, string.Empty) as WMath;

                if (math != null)
                {
                    // Get the current LaTeX string and replace 'x' with 'k'
                    string laTex = math.MathParagraph.LaTeX;
                    math.MathParagraph.LaTeX = laTex.Replace("x", "k");
                }

                // Return the updated document as a downloadable file
                return CreateFileResult(document, "EditEquation.docx");
            }
        }

        // Helper method to create a FileStreamResult for downloading the document
        private FileStreamResult CreateFileResult(WordDocument document, string fileName)
        {
            MemoryStream outputStream = new MemoryStream();
            document.Save(outputStream, FormatType.Docx); // Save document to memory stream
            outputStream.Position = 0; // Reset stream position
            return File(outputStream, "application/docx", fileName); // Return file as response
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
