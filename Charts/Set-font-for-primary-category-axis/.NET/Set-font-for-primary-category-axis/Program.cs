using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;


//Create a new Word document.
using (WordDocument document = new WordDocument(@"../../../Data/Sample.docx", FormatType.Docx))
{
    WChart chart = document.Sections[0].Paragraphs[0].ChildEntities[0] as WChart;
    //Set font to the title of the primary category axis.
    chart.PrimaryCategoryAxis.TitleArea.FontName = "Gigi";
    chart.PrimaryCategoryAxis.TitleArea.Size = 20;

    //Save the Word document to the file stream.
    document.Save(Path.GetFullPath(@"../../../Output/Output.docx"), FormatType.Docx);
}