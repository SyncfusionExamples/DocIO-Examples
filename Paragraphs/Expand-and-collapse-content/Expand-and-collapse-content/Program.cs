using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

// Create a new Word document and add a section to it.
WordDocument document = new WordDocument();
WSection section = document.AddSection() as WSection;

// Add a main heading with Heading 1 style.
WParagraph paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("The Giant Panda");
paragraph.ApplyStyle(BuiltinStyle.Heading1);

// Add a descriptive paragraph following the main heading.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("The giant panda, which only lives in China outside of captivity, has captured the hearts of people of all ages across the globe. From their furry black and white bodies to their shy and docile nature, they are considered one of the world's most loved animals.");

// Add a subheading under the main section and apply Heading 2 style
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Opposable Pseudo Thumb");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a paragraph describing content relevant to the first subheading.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("A characteristic of the giant panda that has mystified scientists is their movable, elongated wrist bone that acts like an opposable thumb. This human-like quality that helps give them even more of a cuddly-panda appearance enables the giant panda to pick up objects and even eat sitting up.");

// Add another subheading under the main section and apply Heading 2 style.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Small panda or Large Raccoon?");
paragraph.ApplyStyle(BuiltinStyle.Heading2);

// Add a paragraph with content relevant to the second subheading.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Giant pandas are generally referred to as bears and are typically called panda bears rather than giant pandas.");

// Add a second main heading and apply Heading 1 style.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Adventure Works Cycles");
paragraph.ApplyStyle(BuiltinStyle.Heading1);

// Add a paragraph with descriptive content for the second main section.
paragraph = section.AddParagraph() as WParagraph;
paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

// Save the document.
using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
{
    document.Save(outputStream1, FormatType.Docx);
}