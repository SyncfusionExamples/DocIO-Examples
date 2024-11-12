using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

// Create a new Word document and add a section to it.
WordDocument document = new WordDocument();
WSection section = document.AddSection() as WSection;

// Add a paragraph and append text to it.
WParagraph paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("The Giant Panda");
// Apply heading level 1.
paragraph.ApplyStyle(BuiltinStyle.Heading1);
// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("The giant panda, which only lives in China outside of captivity, has captured the hearts of people of all ages across the globe.");

// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Small panda or Large Raccoon?");
// Apply heading level 2.
paragraph.ApplyStyle(BuiltinStyle.Heading2);
// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Giant pandas are generally referred to as bears and are typically called panda bears rather than giant pandas.it has several characteristics in common with the red panda.");

// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Adventure Works Cycles");
// Apply heading level 1.
paragraph.ApplyStyle(BuiltinStyle.Heading1);
// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Product Overview");
// Apply heading level 2.
paragraph.ApplyStyle(BuiltinStyle.Heading2);
// Add a paragraph and append text to it.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("While its base operation is located in Bothell, Washington with 290 employees, several regional sales teams are located throughout their market base.");

// Save the document.
using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
{
    document.Save(outputStream, FormatType.Docx);
}