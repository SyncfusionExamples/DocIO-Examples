using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Replace_logo_and_image_with_text
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Open the logo image stream to compare with existing images in the document.
                    FileStream logoImageStream = new FileStream(@"Data/LogoImage.jpg", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    //Find all pictures in the Word document by their EntityType
                    List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);
                    //Iterate through each picture found in the document.
                    foreach (WPicture picture in pictures)
                    {
                        WParagraph ownerParagraph = picture.OwnerParagraph;
                        //Find the index of picture in the owner paragraph.
                        int index = ownerParagraph.ChildEntities.IndexOf(picture);
                        //Remove the picture from its owner paragraph.
                        ownerParagraph.ChildEntities.Remove(picture);
                        //Create a new text range for inserting text.
                        WTextRange textRange = new WTextRange(document);
                        //Check if the current picture's image byte length matches the logo image's length.
                        if (picture.ImageBytes.Length == logoImageStream.Length)
                            //Replace the logo image with the text "< Logo image here >".
                            textRange.Text = "< Logo image here >";
                        else
                            //Replace other images with the text "< Product image here >".
                            textRange.Text = "< Product image here >";
                        //Insert the newly created text range at the index of the removed picture.
                        ownerParagraph.ChildEntities.Insert(index, textRange);
                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
