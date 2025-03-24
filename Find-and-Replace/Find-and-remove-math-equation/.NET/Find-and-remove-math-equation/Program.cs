using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Find_and_remove_math_equation
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                   //Find all Equations by EntityType in Word document.
                    List<Entity> equations = document.FindAllItemsByProperty(EntityType.Math, null, null);
                    //Remove the equation.
                    for (int i = 0; i < equations.Count; i++)
                    {
                        WMath equation = equations[i] as WMath;
                        equation.OwnerParagraph.ChildEntities.Remove(equation);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
