using Syncfusion.DocIO.DLS;
using System.IO;

//Loads an existing Word document into DocIO instance.
WordDocument doc = new WordDocument(@"../../../Data/Template.docx");

WParagraph firstPara = doc.Sections[0].Paragraphs[0];
WParagraph temp = new WParagraph(doc);
WPicture pic = temp.AppendPicture(File.ReadAllBytes(@"../../../Data/Image.png")) as WPicture;
pic.Width = 400;
pic.Height = 200;
firstPara.Items.Insert(0, temp.Items[0]);

//Saving the document as .docx
doc.Save(@"../../../Output/Result.docx");
//Closes the document instance.
doc.Close();