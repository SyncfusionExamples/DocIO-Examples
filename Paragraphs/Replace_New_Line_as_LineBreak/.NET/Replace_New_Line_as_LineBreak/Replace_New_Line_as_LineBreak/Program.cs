using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_New_Line_as_LineBreak
{
    class Program
    {
        static void Main(string[] args)
        {
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("NxYtFisQPR08Cit/VkR+XU9FfV5AQmBIYVp/TGpJfl96cVxMZVVBJAtUQF1hTH9SdENiWHtZc3ZVRWFeWkd1");
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //This method adds a section and a paragraph in the document.
                document.EnsureMinimal();
                //Accessing the last paragraph.
                IWParagraph paragraph = document.LastParagraph;
                //Applying bullet style to the paragraph.
                paragraph.ListFormat.ApplyDefBulletStyle();
                //Text to be added.
                string completeText = "Adventure Works Cycles:\nAdventure Works Cycles, the fictitious company on which the Adventure Works sample databases are based, is a large, multinational manufacturing company.\n" +
                    "The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.";
                //Splitting the text based on new line character.
                string[] textArray = completeText.Split('\n');
                //Adding the splitted text in the paragraph by inserting the LineBreak instead of new line character.
                foreach (string text in textArray)
                {
                    paragraph.AppendText(text);
                    paragraph.AppendBreak(BreakType.LineBreak);
                }
                //Saving the document.
                using (FileStream outputStream = new FileStream(@"../../../Output/Result.docx", FileMode.OpenOrCreate))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
        