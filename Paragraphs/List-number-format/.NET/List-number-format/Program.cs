using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;


// Creates a new instance of WordDocument to work with.
using (WordDocument document = new WordDocument())
{
    // Adds a new section to the document.
    IWSection section = document.AddSection();
    IWParagraph paragraph = section.AddParagraph();

    // Adds a numbered list style with CardinalText pattern (One, Two, Three, ...).
    ListStyle listStyle = document.AddListStyle(ListType.Numbered, "CardinalText");
    WListLevel levelOne = listStyle.Levels[0];
    levelOne.PatternType = ListPatternType.CardinalText;
    levelOne.StartAt = 1;

    // Adds a heading paragraph for the CardinalText list.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List pattern Cardinal Text");

    // Adds first list item using CardinalText style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 1");
    paragraph.ListFormat.ApplyStyle("CardinalText");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds second list item using CardinalText style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 2");
    paragraph.ListFormat.ApplyStyle("CardinalText");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds third list item using CardinalText style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 3");
    paragraph.ListFormat.ApplyStyle("CardinalText");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds a blank paragraph before the next list.
    paragraph = section.AddParagraph();

    // Adds a numbered list style with HindiLetter1 pattern.
    listStyle = document.AddListStyle(ListType.Numbered, "HindiLetter1");
    levelOne = listStyle.Levels[0];
    levelOne.PatternType = ListPatternType.HindiLetter1;
    levelOne.StartAt = 1;

    // Adds a heading paragraph for the HindiLetter1 list.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List pattern Hindi Letter");

    // Adds first list item using HindiLetter1 style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 1");
    paragraph.ListFormat.ApplyStyle("HindiLetter1");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds second list item using HindiLetter1 style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 2");
    paragraph.ListFormat.ApplyStyle("HindiLetter1");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds third list item using HindiLetter1 style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 3");
    paragraph.ListFormat.ApplyStyle("HindiLetter1");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds a blank paragraph before the next list.
    paragraph = section.AddParagraph();

    // Adds a numbered list style with Hebrew1 pattern.
    listStyle = document.AddListStyle(ListType.Numbered, "Hebrew1");
    levelOne = listStyle.Levels[0];
    levelOne.PatternType = ListPatternType.Hebrew1;
    levelOne.StartAt = 1;

    // Adds a heading paragraph for the Hebrew1 list.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List pattern Herbrew");

    // Adds first list item using Hebrew1 style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 1");
    paragraph.ListFormat.ApplyStyle("Hebrew1");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds second list item using Hebrew1 style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 2");
    paragraph.ListFormat.ApplyStyle("Hebrew1");
    paragraph.ListFormat.ContinueListNumbering();

    // Adds third list item using Hebrew1 style.
    paragraph = section.AddParagraph();
    paragraph.AppendText("List item 3");
    paragraph.ListFormat.ApplyStyle("Hebrew1");
    paragraph.ListFormat.ContinueListNumbering();

    // Create a file stream to save the document.
    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
    {
        // Saves the Word document to the specified file stream in DOCX format.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
