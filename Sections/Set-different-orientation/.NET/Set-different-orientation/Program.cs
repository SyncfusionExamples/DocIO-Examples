using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open an existing document
using (WordDocument document = new WordDocument(new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read), FormatType.Docx))
{
    //Set different orientations
    document.Sections[0].PageSetup.Orientation = PageOrientation.Portrait;
    document.Sections[1].PageSetup.Orientation = PageOrientation.Landscape;
    document.Sections[2].PageSetup.Orientation = PageOrientation.Portrait;
    document.Sections[3].PageSetup.Orientation = PageOrientation.Landscape;

    //Save the Word document
    document.Save(new FileStream(@"../../../Output.docx", FileMode.Create, FileAccess.Write), FormatType.Docx);
}