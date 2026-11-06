using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Loads an existing Word document into DocIO instance.
using (WordDocument document = new WordDocument(@"../../../Data/Template.docx"))
{
    //Access form fields by name.
    (document.LastSection.Body.FormFields["FullName"] as WTextFormField).TextRange.Text = "John";
    (document.LastSection.Body.FormFields["BirthDayField"] as WTextFormField).TextRange.Text = "5.13.1980";

    //Access form fields by index.
    (document.LastSection.Body.FormFields[2] as WTextFormField).TextRange.Text = "221b Baker Street";
    (document.LastSection.Body.FormFields[3] as WTextFormField).TextRange.Text = "(206)555-3412";
    (document.LastSection.Body.FormFields[4] as WTextFormField).TextRange.Text = "John@company.com";
    (document.LastSection.Body.FormFields[5] as WDropDownFormField).DropDownSelectedIndex = 1;
    (document.LastSection.Body.FormFields[6] as WTextFormField).TextRange.Text = "Michigan University";
    (document.LastSection.Body.FormFields[7] as WDropDownFormField).DropDownSelectedIndex = 1;
    (document.LastSection.Body.FormFields[8] as WDropDownFormField).DropDownSelectedIndex = 2;

    document.Save(@"../../../Output/Sample.docx", FormatType.Docx);
}
