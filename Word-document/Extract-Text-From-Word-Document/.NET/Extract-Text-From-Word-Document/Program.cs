using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Loads a template document
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Get the document text
    string text = document.GetText();

    // Save the text to a file
    string outputPath = Path.GetFullPath(@"Output/Output.txt");
    File.WriteAllText(outputPath, text);
}
