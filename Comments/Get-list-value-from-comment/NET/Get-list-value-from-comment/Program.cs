using System;
using System.IO;
using System.Collections.Generic;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
namespace Get_list_value_from_comment
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Gets the document text
                    document.GetText();
                    foreach (WComment comment in document.Comments)
                    {
                        List<string> str = new List<string>();
                        for (int y = 0; y < comment.ChildEntities.Count; y++)
                        {
                            WParagraph paragraph = (Syncfusion.DocIO.DLS.WParagraph)comment.ChildEntities[y];
                            if (paragraph.ListFormat != null && paragraph.ListFormat.CurrentListLevel != null)
                                str.Add(paragraph.ListString + " " + paragraph.Text);
                        }
                        if (str.Count > 0)
                            Console.WriteLine($"{String.Join("\n", str)}");
                    }
                }
            }
        }
    }
}
