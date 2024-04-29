using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream fileStreamPath = new FileStream(@"../../../Data/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open the existing Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Change section formatting
        ChangeSectionFormatting(document);
        //Change paragraph style formatting
        ChangeParagraphFormatting(document);
        //Change table style formatting
        ChangeTableFormatting(document);
        using (FileStream outputStream = new FileStream(@"../../../Data/Output.docx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            //Save the Word document.
            document.Save(outputStream, FormatType.Docx);
        }
    }
}

//Modify the section formatting like page orientation.
void ChangeSectionFormatting(WordDocument document)
{
    //Iterate each section.
    for (int i = 0; i < document.Sections.Count; i++)
    {
        //Set the Orientation.
        document.Sections[i].PageSetup.Orientation = PageOrientation.Landscape;
        //Set the top margin.
        document.Sections[i].PageSetup.Margins.Top = 100;
    }
}

//Word document uses "Normal" paragraph style as default for all paragraphs.
//Modify the default paragraph format "Normal"
void ChangeParagraphFormatting(WordDocument document)
{
    //Get the default paragraph style.
    WParagraphStyle paraStyle = document.Styles.FindByName("Normal") as WParagraphStyle;
    //Set character format.
    paraStyle.CharacterFormat.FontName = "Arial";
    paraStyle.CharacterFormat.FontSize = 14;
    //Set paragraph format.
    paraStyle.ParagraphFormat.AfterSpacing = 20;
}

//Word document uses "Table Grid" table style as default for all tables.
//Modify the deafult table format "Table Grid"
void ChangeTableFormatting(WordDocument document)
{
    //Get the default table style.
    WTableStyle tableStyle = document.Styles.FindByName("Table Grid") as WTableStyle;
    //Set cell spacing to the table.
    tableStyle.TableProperties.CellSpacing = 5;
    //Applied BackColor to the table.
    tableStyle.TableProperties.BackColor = Syncfusion.Drawing.Color.Blue;
}