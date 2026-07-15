using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Loads a template document
WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx"));
// Gets the document text
string text = document.GetText();
// Prints the extracted text to the console
Console.WriteLine(text);
// Dispose the document instance 
document.Close();
