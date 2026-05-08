using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Get_Type_and_Name_of_ContentControls
{
    class Program
    {

        static void Main(string[] args)
        {
            using (FileStream docStream = new FileStream(@Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the file stream into a Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    //Find all block content controls
                    List<Entity> blockContentControls = document.FindAllItemsByProperties(EntityType.BlockContentControl, null, null);

                    //Process block content controls
                    foreach (BlockContentControl blockContentControl in blockContentControls)
                    {
                        //Access block content control properties
                        string title = blockContentControl.ContentControlProperties.Title;
                        string tag = blockContentControl.ContentControlProperties.Tag;
                    }

                    //Find all inline content controls
                    List<Entity> inlineContentControls = document.FindAllItemsByProperties(EntityType.InlineContentControl, null, null);

                    //Process inline content controls
                    foreach (InlineContentControl inlineContentControl in inlineContentControls)
                    {
                        //Access inline content control properties
                        string title = inlineContentControl.ContentControlProperties.Title;
                        string tag = inlineContentControl.ContentControlProperties.Tag;
                    }

                    //Save the modified document if needed
                    using (FileStream outputStream = new FileStream(@Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }

        }
    }
}
