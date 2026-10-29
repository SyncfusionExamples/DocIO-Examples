using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Loads an existing Word document into DocIO instance.
using (WordDocument document = new WordDocument(@"../../../Data/Template.docx"))
{
    IWSection section = document.LastSection;

    (section.Body.FormFields["FullName"] as WTextFormField).TextRange.Text = "John";
    (section.Body.FormFields["BirthDayField"] as WTextFormField).TextRange.Text = "5.13.1980";

    (section.Body.FormFields[2] as WTextFormField).TextRange.Text = "221b Baker Street";
    (section.Body.FormFields[3] as WTextFormField).TextRange.Text = "(206)555-3412";
    (section.Body.FormFields[4] as WTextFormField).TextRange.Text = "John@company.com";

    (section.Body.FormFields[5] as WDropDownFormField).DropDownSelectedIndex = 1;
    (section.Body.FormFields[6] as WTextFormField).TextRange.Text = "Michigan University";
    (section.Body.FormFields[7] as WDropDownFormField).DropDownSelectedIndex = 1;
    (section.Body.FormFields[8] as WDropDownFormField).DropDownSelectedIndex = 2;

    document.Save(@"../../../Output/Sample.docx", FormatType.Docx);
}