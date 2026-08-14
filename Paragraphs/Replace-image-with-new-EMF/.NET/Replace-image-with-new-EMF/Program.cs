using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

using (WordDocument document = new WordDocument(Path.Combine(@"../../../Data/input.docx"), FormatType.Docx))
{
    //Find picture in document
    List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);

    if (pictures != null)
    {
        //Iterate and replace image
        foreach (Entity entity in pictures)
        {
            WPicture picture = entity as WPicture;
            FileStream imageStream = new FileStream(@"../../../Data/Image.emf", FileMode.Open, FileAccess.ReadWrite);
            picture.LoadImage(imageStream);
            imageStream.Close();

        }
    }
    document.Save(@"../../../Output/output.docx", FormatType.Docx);
}