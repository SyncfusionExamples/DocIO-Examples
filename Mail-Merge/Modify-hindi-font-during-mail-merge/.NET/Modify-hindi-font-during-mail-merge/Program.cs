using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_hindi_font_during_mail_merge
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
                    string[] fieldNames = new string[] { "EmployeeId", "Name", "Phone", "City" };
                    string[] fieldValues = new string[] { "1001", "नैन्सी डेवियलो", "+122-2222222", "London" };

                    // Uses the mail merge events to perform the conditional formatting during runtime.
                    document.MailMerge.MergeField += new MergeFieldEventHandler(ApplyFontForHindiText);
                    //Performs the mail merge.
                    document.MailMerge.Execute(fieldNames, fieldValues);

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
        /// Represents the method that handles the MergeField event.
        /// </summary>      
        private static void ApplyFontForHindiText(object sender, MergeFieldEventArgs args)
        {
            string fieldValue = args.FieldValue.ToString();
            //If the field value contains Hindi characters, then apply the font.
            bool containsHindi = ContainsHindiCharacters(fieldValue);
            if (containsHindi)
            {
                args.TextRange.CharacterFormat.FontName = "Nirmala UI";
            }
        }

        /// <summary>
        /// Checks whether the given character is a Hindi character or not.
        /// </summary>
        private static bool IsHindiChar(char character)
        {
            //Hindi characters are comes under the Devanagari scripts.
            //The Unicode Standard defines three blocks for Devanagari. https://en.wikipedia.org/wiki/Devanagari#Unicode              
            return ((character >= '\u0900' && character <= '\u097f') //Devanagari (U+0900–U+097F), https://en.wikipedia.org/wiki/Devanagari_(Unicode_block)
                || (character >= '\ua8e0' && character <= '\ua8ff') //Devanagari Extended (U+A8E0–U+A8FF), https://en.wikipedia.org/wiki/Devanagari_Extended
                || (character >= '\u1cd0' && character <= '\u1cff')); //Vedic Extensions (U+1CD0–U+1CFF), https://en.wikipedia.org/wiki/Vedic_Extensions
        }
        /// <summary>
        /// Checks whether the given text contains Hindi characters or not.
        /// </summary>
        private static bool ContainsHindiCharacters(string text)
        {
            foreach (char character in text)
            {
                if (IsHindiChar(character))
                {
                    return true; // If any Hindi character is found, return true.
                }
            }
            return false; // No Hindi characters were found.
        }
	}
}
