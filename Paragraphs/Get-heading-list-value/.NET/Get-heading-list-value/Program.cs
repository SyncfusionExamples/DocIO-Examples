using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System;
using System.IO;

namespace Get_heading_list_value
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document from the file stream.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Get the document text.
                    document.GetText();
                    //Find all paragraphs with the style 'Heading 3' in the Word document.
                    List<Entity> headingParagraphs = document.FindAllItemsByProperty(EntityType.Paragraph, "StyleName", "Heading 3");
                    if (headingParagraphs.Count == 0)
                        Console.WriteLine("No paragraphs with the style 'Heading 3' found.");
                    else
                    {
                        foreach (Entity paragraph in headingParagraphs)
                        {
                            //Get the string that represents the appearance of list value of the paragraph.
                            if (paragraph is WParagraph)
                                Console.WriteLine((paragraph as WParagraph).ListString);
                            else
                                Console.WriteLine("The entity is not a WParagraph.");
                        }
                    }
                    //Pauses the console to display the output.
                    Console.ReadLine();
                }
            }
        }
    }
}
