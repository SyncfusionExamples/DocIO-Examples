using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Extract_images_from_Word_document
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    int index = 0;
                    //Visits all document entities.
                    foreach (var item in document.Visit())
                    {
                        switch (item.EntityType)
                        {
                            case EntityType.Picture:
                                WPicture picture = item as WPicture;

                                // Use a MemoryStream to handle the image bytes from the picture
                                using (MemoryStream memoryStream = new MemoryStream(picture.ImageBytes))
                                {
                                    // Define the path where the image will be saved
                                    string imagePath = Path.GetFullPath(@"../../../Images/Image-" + index + ".jpeg");

                                    // Create a FileStream to write the image to the specified path
                                    using (FileStream image = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                                    {
                                        // Copy the content of the MemoryStream to the FileStream
                                        memoryStream.CopyTo(image);
                                    }
                                }

                                // Increment the index for the next image
                                index++;
                                break;
                        }

                    }
                }
            }
        }
    }

    public static class DocIOExtensions
    {
        public static IEnumerable<IEntity> Visit(this ICompositeEntity entity)
        {
            var entities = new Stack<IEntity>(new IEntity[] { entity });
            while (entities.Count > 0)
            {
                var e = entities.Pop();
                yield return e;
                if (e is ICompositeEntity)
                {
                    foreach (IEntity childEntity in ((ICompositeEntity)e).ChildEntities)
                    {
                        entities.Push(childEntity);
                    }
                }
            }
        }
    }
}
