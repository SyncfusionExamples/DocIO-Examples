using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open an existing document
using (FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        //Set different orientations
        document.Sections[0].PageSetup.Orientation = PageOrientation.Portrait;
        document.Sections[1].PageSetup.Orientation = PageOrientation.Landscape;
        document.Sections[2].PageSetup.Orientation = PageOrientation.Portrait;
        document.Sections[3].PageSetup.Orientation = PageOrientation.Landscape;

        //Save the Word document
        using (FileStream outputStream = new FileStream(@"Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}