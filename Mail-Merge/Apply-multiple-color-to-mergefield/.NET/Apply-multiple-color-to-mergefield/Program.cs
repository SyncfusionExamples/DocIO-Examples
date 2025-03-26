using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_multiple_color_to_mergefield
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    string[] fieldNames = new string[] { "RedBlack", "RedBlackGreen" };
                    string[] fieldValues = new string[] { "Red Black", "Red Black Green" };
                    //Creates mail merge events handler to split the field value and applies the color
                    document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
                    //Performs the mail merge
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    //Removes mail merge events handler
                    document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);

                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Splits the field value and applies the color by using MergeFieldEventHandler.
        /// </summary>
        public static void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            if (args.FieldName == "RedBlack" || args.FieldName == "RedBlackGreen")
            {
                //Split the field result value based on space between the words.
                string[] splitText = args.FieldValue.ToString().Split(' ');

                //Modifies the field result text as "Red" and applies the color red.
                args.TextRange.Text = splitText[0];
                if (args.TextRange.Text == "Red")
                    args.TextRange.CharacterFormat.TextColor = Color.Red;

                //Gets the merge field owner paragraph.
                WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                //Gets the index of merge field
                int fieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                //Gets the index next to the merge field.
                int fieldNextIndex = GetFieldNextIndex(fieldIndex, paragraph);

                //Appends the remaining texts after the merge field and applies the color.
                for (int i = 1; i < splitText.Length; i++)
                {
                    //Initialize new text range.
                    WTextRange textRange = new WTextRange(paragraph.Document);
                    //Specifies the text.
                    textRange.Text = " " + splitText[i];
                    //Applies the color based on the text
                    if (textRange.Text == " " + "Black")
                        textRange.CharacterFormat.TextColor = Color.Black;
                    else if (textRange.Text == " " + "Green")
                        textRange.CharacterFormat.TextColor = Color.Green;

                    //Appends the text range after the merge field.
                    if (fieldNextIndex != -1 && fieldNextIndex < paragraph.ChildEntities.Count)
                        paragraph.ChildEntities.Insert(fieldNextIndex, textRange);
                    else
                        paragraph.ChildEntities.Add(textRange);
                    fieldNextIndex++;
                }
            }
        }
        /// <summary>
        /// Returns the index next to the merge field.
        /// </summary>
        public static int GetFieldNextIndex(int fieldIndex, WParagraph paragraph)
        {
            for (int i = fieldIndex; i < paragraph.ChildEntities.Count; i++)
            {
                ParagraphItem item = paragraph.ChildEntities[i] as ParagraphItem;
                if (item != null && item is WFieldMark && (item as WFieldMark).Type == FieldMarkType.FieldEnd)
                    return i + 1;
            }
            return -1;
        }
    }
}
