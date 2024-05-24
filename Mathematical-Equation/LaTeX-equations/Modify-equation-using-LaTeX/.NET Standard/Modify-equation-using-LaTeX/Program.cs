using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Open a Word document using File stream.
using (FileStream inputStream = new FileStream("../../../Input.docx", FileMode.Open, FileAccess.Read))
{
    // OPen the existing Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Access the first paragraph from the last section of the document
        WParagraph paragraph = document.LastSection.Body.ChildEntities[2] as WParagraph;

        // Retrieve the first math equation in the paragraph, if it exists
        WMath math = paragraph.ChildEntities[0] as WMath;
        if (math != null)
        {
            // Get the LaTeX representation of the math equation
            string laTeX = math.MathParagraph.LaTeX;
            // Replace occurrences of 'a' with 's' in the LaTeX representation
            laTeX = laTeX.Replace("x", "k");
            //Modify the LaTeX string
            math.MathParagraph.LaTeX = laTeX;
        }
        using (FileStream outputStream = new FileStream(@"../../../Result.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}