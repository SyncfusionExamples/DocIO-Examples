using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Open a Word document using File stream.
using (FileStream inputStream = new FileStream("../../../Input.docx", FileMode.Open, FileAccess.Read))
{
    // OPen the existing Word document.
    using (WordDocument document = new WordDocument())
    {
        List<Entity> entities = document.FindAllItemsByProperty(EntityType.Math, null, null);

        //Iterate through each equation in the Word document.
        foreach (Entity entity in entities)
        {
            WMath math = entity as WMath;
            //Get the laTeX code of equation.
            string laTeX = math.MathParagraph.LaTeX;
            //Print the LaTeX equation
            Console.WriteLine(laTeX + "\n");
        }
    }
}