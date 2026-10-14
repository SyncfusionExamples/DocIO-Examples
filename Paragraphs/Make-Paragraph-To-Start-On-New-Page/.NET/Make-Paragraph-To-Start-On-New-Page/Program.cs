using Syncfusion.DocIO.DLS;
using System.IO;

namespace Make_Paragraph_To_Start_On_New_Page
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document instance.
            using (WordDocument document = new WordDocument())
            {
                //Adds the section into Word document.
                IWSection section = document.AddSection();

                //Adds a paragraph to created section.
                IWParagraph firstPageParagraph = section.AddParagraph();

                //Appends the text to the created paragraph.
                IWTextRange textRange = firstPageParagraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

                //Adds a paragraph move to the new page 
                IWParagraph paragraph = section.AddParagraph();

                //Sets Page break.
                paragraph.ParagraphFormat.PageBreakBefore = true;

                //Appends the text to the created paragraph.
                IWTextRange newPageTextRange = paragraph.AppendText("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");

                document.Save(Path.GetFullPath("Output/Output.docx"));
            }
        }
    }
}
