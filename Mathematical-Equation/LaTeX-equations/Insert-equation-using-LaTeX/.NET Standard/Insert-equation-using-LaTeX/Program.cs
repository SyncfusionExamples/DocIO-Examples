using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Open a Word document using File stream.
using (FileStream inputStream = new FileStream("../../../Input.docx", FileMode.Open, FileAccess.Read)) {
    // OPen the existing Word document.
    using (WordDocument document = new WordDocument())
    {
        //Create a new mathematical equation.
        WMath math = new WMath(document);
        //Set the LaTeX string to the equation .
        math.MathParagraph.LaTeX = @"f\left(x\right)={a}_{0}+\sum_{n=1}^{\infty}{\left({a}_{n}\cos{\frac{n\pi{x}}{L}}+{b}_{n}\sin{\frac{n\pi{x}}{L}}\right)}";
        //Insert the new mathematical equation to existing paragraph.
        document.LastSection.Paragraphs[2].ChildEntities.Insert(0, math);

        //Save the Word document.
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    } 
}