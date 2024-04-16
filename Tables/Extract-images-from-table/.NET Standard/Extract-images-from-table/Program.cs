using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


//Creates an instance of WordDocument class
FileStream fileStreamPath = new FileStream("../../../Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);

//Get the list of pictures
List<Entity> entities = document.FindAllItemsByProperty(EntityType.Picture, "OwnerParagraph.IsInCell", "True");

//Closes the document
document.Close();