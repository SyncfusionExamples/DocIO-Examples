using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_and_replace_with_merge_field
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Finds all the placeholder text enclosed within '«' and '»' in the Word document.
                    TextSelection[] textSelections = document.FindAll(new Regex("«([(?i)image(?-i)]*:*[a-zA-Z0-9 ]*:*[a-zA-Z0-9 ]+)»"));
                    string[] searchedPlaceholders = new string[textSelections.Length];
                    for (int i = 0; i < textSelections.Length; i++)
                    {
                        searchedPlaceholders[i] = textSelections[i].SelectedText;
                    }
                    for (int i = 0; i < searchedPlaceholders.Length; i++)
                    {
                        //Replaces the placeholder text enclosed within '«' and '»' with desired merge field.
                        WParagraph paragraph = new WParagraph(document);
                        paragraph.AppendField(searchedPlaceholders[i].TrimStart('«').TrimEnd('»'), FieldType.FieldMergeField);
                        TextSelection newSelection = new TextSelection(paragraph, 0, paragraph.Items.Count);
                        TextBodyPart bodyPart = new TextBodyPart(document);
                        bodyPart.BodyItems.Add(paragraph);
                        document.Replace(searchedPlaceholders[i], bodyPart, true, true, true);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
