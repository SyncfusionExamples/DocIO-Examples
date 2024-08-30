using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_footnotes_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
				//Creates a section.
				IWSection section = document.AddSection();
				//Adds a paragraph to a section.
				IWParagraph paragraph = section.AddParagraph();
				//Appends the text to paragraph.
				paragraph.AppendText("Working with footnotes");
				//Formats the text.
				paragraph.ApplyStyle(BuiltinStyle.Heading1);
				//Adds a paragraph to a section.
				paragraph = section.AddParagraph();
				//Appends the footnotes.
				WFootnote footnote = (WFootnote)paragraph.AppendFootnote(Syncfusion.DocIO.FootnoteType.Footnote);
				//Sets the footnote character format.
				footnote.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
				//Inserts the text into the paragraph.
				paragraph.AppendText("Sample content for footnotes").CharacterFormat.Bold = true;
				//Adds footnote text.
				paragraph = footnote.TextBody.AddParagraph();
				paragraph.AppendText("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
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
