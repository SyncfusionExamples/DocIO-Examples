using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open an existing document
using (FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        //Find all instances of the target word in the document
        TextSelection[] textSelections = document.FindAll("Barcode", false, true);

        foreach (TextSelection selection in textSelections)
        {
            //Apply the barcode font style
            selection.GetAsOneRange().CharacterFormat.FontName = "Code 128";
        }

        //Save the Word document
        using (FileStream outputStream = new FileStream(@"Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}