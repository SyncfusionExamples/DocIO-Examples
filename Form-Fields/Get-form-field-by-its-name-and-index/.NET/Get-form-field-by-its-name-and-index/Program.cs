using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

//Loads an existing Word document into DocIO instance.
WordDocument document = new WordDocument(@"../../../Data/Template.docx");
IWSection sec = document.LastSection;
WTextFormField textFF;
WDropDownFormField dropFF;

//Access the form field with field name
textFF = sec.Body.FormFields["FullName"] as WTextFormField;

//Fill value for the textfield
textFF.TextRange.Text = "John";

//Access the form field with field name
textFF = sec.Body.FormFields["BirthDayField"] as WTextFormField;
textFF.TextRange.Text = "5.13.1980";

//Access the form field with index
textFF = sec.Body.FormFields[2] as WTextFormField;
textFF.TextRange.Text = "221b Baker Street";

textFF = sec.Body.FormFields[3] as WTextFormField;
textFF.TextRange.Text = "(206)555-3412";

textFF = sec.Body.FormFields[4] as WTextFormField;
textFF.TextRange.Text = "John@company.com";

dropFF = sec.Body.FormFields[5] as WDropDownFormField;

//Set the value
dropFF.DropDownSelectedIndex = 1;

textFF = sec.Body.FormFields[6] as WTextFormField;
textFF.TextRange.Text = "Michigan University";

dropFF = sec.Body.FormFields[7] as WDropDownFormField;
dropFF.DropDownSelectedIndex = 1;

dropFF = sec.Body.FormFields[8] as WDropDownFormField;
dropFF.DropDownSelectedIndex = 2;

//Saving the document as .docx
document.Save(@"../../../Output/Sample.docx", FormatType.Docx);
//Closes the document instance.
document.Close();