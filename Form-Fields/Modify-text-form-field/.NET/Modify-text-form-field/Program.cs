using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_text_form_field
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Iterates through section.
                    foreach (WSection section in document.Sections)
                    {
                        //Iterates through section child elements.
                        foreach (WTextBody textBody in section.ChildEntities)
                        {
                            //Iterates through form fields.
                            foreach (WFormField formField in textBody.FormFields)
                            {
                                switch (formField.FormFieldType)
                                {
                                    case FormFieldType.TextInput:
                                        WTextFormField textField = formField as WTextFormField;
                                        if (textField.Type == TextFormFieldType.DateText)
                                        {
                                            //Modifies the text form field.
                                            textField.Type = TextFormFieldType.RegularText;
                                            textField.StringFormat = "";
                                            textField.DefaultText = "Default text";
                                            textField.Text = "Default text";
                                            textField.CalculateOnExit = false;
                                        }
                                        break;
                                }
                            }
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
