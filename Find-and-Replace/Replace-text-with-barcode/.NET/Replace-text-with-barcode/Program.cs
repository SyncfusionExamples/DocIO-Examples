using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Replace_text_with_barcode
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open an existing document
            using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read))
            {
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Find all instances of the target word in the document
                    TextSelection[] textSelections = document.FindAll("Barcode", false, true);

                    foreach (TextSelection selection in textSelections)
                    {
                        //Apply the font style
                        selection.GetAsOneRange().CharacterFormat.FontName = "Code 128";
                        selection.GetAsOneRange().CharacterFormat.Bold = true;
                        selection.GetAsOneRange().CharacterFormat.Italic = true;
                    }

                    //Save the Word document
                    using (FileStream outputStream = new FileStream(@"../../../Output.docx", FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
