using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Ungroup_shapes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the last paragraph.
                    WParagraph lastParagraph = document.LastParagraph;
                    //Iterates through the paragraph items to get the group shape.
                    for (int i = 0; i < lastParagraph.ChildEntities.Count; i++)
                    {
                        if (lastParagraph.ChildEntities[i] is GroupShape)
                        {
                            GroupShape groupShape = lastParagraph.ChildEntities[i] as GroupShape;
                            //Ungroup the child shapes in the group shape.
                            groupShape.Ungroup();
                            break;
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
