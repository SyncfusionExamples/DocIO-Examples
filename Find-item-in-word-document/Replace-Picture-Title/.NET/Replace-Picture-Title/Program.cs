using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_Picture_Title
{
    class Program
    {

        static void Main(string[] args)
        {
            //Load an existing Word document.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument wordDocument = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Find all pictures with the title "Picture1" in the Word document.
                    Entity picture = wordDocument.FindItemByProperty(EntityType.Picture, "Title", "Bookmark");
                    if (picture != null)
                    {
                        WPicture wPicture = picture as WPicture;
                        wPicture.Title = "Red_Handed Cycle";
                    }
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Output/Result.docx"), FileMode.Create))
                    {
                        wordDocument.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
