using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


//Creates an instance of WordDocument class
FileStream fileStreamPath = new FileStream("../../../Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);

//Get the list of pictures
List<Entity> entities = document.FindAllItemsByProperty(EntityType.Picture, "OwnerParagraph.IsInCell", "True");

//Save all the images to a folder
for (int i = 0; i < entities.Count; i++)
{
    WPicture picture = entities[i] as WPicture;
    //Get the image 
    System.Drawing.Image image = System.Drawing.Image.FromStream(new MemoryStream(picture.ImageBytes));
    //Save the image as PNG
    string imgFileName = @"../../../Output" + i + ".png";
    FileStream imgFile = new FileStream(imgFileName, FileMode.Create, FileAccess.ReadWrite);
    image.Save(imgFile, System.Drawing.Imaging.ImageFormat.Png);
    //Dispose the instances.
    imgFile.Dispose();
    image.Dispose();
}

//Closes the document
document.Close();