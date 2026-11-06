using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (WordDocument document = new WordDocument(@"../../../Data/Template.docx"))
{
    FormFieldCollection formFields = document.LastSection.Body.FormFields;

    (formFields["FullName"] as WTextFormField).TextRange.Text = "John";
    (formFields["BirthDayField"] as WTextFormField).TextRange.Text = "5.13.1980";

    ((WTextFormField)formFields[2]).TextRange.Text = "221b Baker Street";
    ((WTextFormField)formFields[3]).TextRange.Text = "(206)555-3412";
    ((WTextFormField)formFields[4]).TextRange.Text = "John@company.com";

    ((WDropDownFormField)formFields[5]).DropDownSelectedIndex = 1;
    ((WTextFormField)formFields[6]).TextRange.Text = "Michigan University";
    ((WDropDownFormField)formFields[7]).DropDownSelectedIndex = 1;
    ((WDropDownFormField)formFields[8]).DropDownSelectedIndex = 2;

    document.Save(@"../../../Output/Sample.docx", FormatType.Docx);
}