using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Extract_text_in_plain_content_control
{
    class Program
    {
        static List<string> extractedTextCollection = new List<string>();
        static void Main(string[] args)
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Creates an instance of WordDocument class
                using (WordDocument document = new WordDocument(docStream, FormatType.Automatic))
                {
                    //Find all plain text content control by EntityType in Word document.
                    List<Entity> plainTextContentControls = document.FindAllItemsByProperty(EntityType.InlineContentControl, "ContentControlProperties.Type", "Text");
                    // Extract text from all plain text content controls.
                    for (int i = 0; i < plainTextContentControls.Count; i++)
                    {
                        InlineContentControl plainTextContentControl = plainTextContentControls[i] as InlineContentControl;
                        ExtractTextInPlainTextContentControl(plainTextContentControl);
                    }
                    // Print extracted text to the console
                    Console.WriteLine("Extracted Text from Plain Text Content Controls:");
                    foreach (string extractedText in extractedTextCollection)
                        Console.WriteLine(extractedText);
                    Console.ReadLine();
                }
            }
        }
        
        /// <summary>
        /// Extract Text in Plain text content control
        /// </summary>
        private static void ExtractTextInPlainTextContentControl(InlineContentControl plainTextContentControl)
        {
            foreach (ParagraphItem item in plainTextContentControl.ParagraphItems)
            {
                if (item is WTextRange)
                {
                    WTextRange textRange = (item as WTextRange);
                    string text = textRange.Text;
                    extractedTextCollection.Add(text);
                }
            }
        }
    }
}
