using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open an existing document
using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        //Get the section
        WSection section = document.Sections[2];

        //Remove the first page header
        section.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
        section.HeadersFooters.FirstPageHeader.AddParagraph();
        //Remove the first page footer
        section.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
        section.HeadersFooters.FirstPageFooter.AddParagraph();

        //Remove the odd header
        section.HeadersFooters.OddHeader.ChildEntities.Clear();
        section.HeadersFooters.OddHeader.AddParagraph();
        //Remove the odd footer
        section.HeadersFooters.OddFooter.ChildEntities.Clear();
        section.HeadersFooters.OddFooter.AddParagraph();

        //Remove the even header
        section.HeadersFooters.EvenHeader.ChildEntities.Clear();
        section.HeadersFooters.EvenHeader.AddParagraph();
        //Remove the even footer
        section.HeadersFooters.EvenFooter.ChildEntities.Clear();
        section.HeadersFooters.EvenFooter.AddParagraph();

        //Save the Word document
        using (FileStream outputStream = new FileStream(@"../../../Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}