using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Extract_Selected_CheckBox
{
    class Program
    {

        static void Main(string[] args)
        {
            List<string> questionList = new List<string>();
            //Load an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../../Data/Template.docx"), FormatType.Docx))
            {
                //Find all checkBoxes in the Word document.
                List<Entity> checkBoxes = wordDocument.FindAllItemsByProperty(EntityType.CheckBox, "Checked", "True");


                if (checkBoxes != null)
                {

                    foreach (WCheckBox checkBox in checkBoxes)
                    {
                        string paragraphText = "";
                        WParagraph ownerParagraph = (checkBox.Owner as WParagraph);
                        int index = ownerParagraph.ChildEntities.IndexOf(checkBox);
                        //Extract the question information when checkBox is the first item of the paragraph.
                        if (index == 0)
                        {
                            //Get the paragraph's item collection.
                            ParagraphItemCollection paraItems = ownerParagraph.Items;
                            // This boolean is used to avoid the checkBox in the extracted question information.
                            bool isIgnoreCheckBoxItems = true;
                            for (int i = 0; i < paraItems.Count; i++)
                            {
                                //Collect the question information from the paragraph.
                                if (!isIgnoreCheckBoxItems)
                                {
                                    switch (paraItems[i].EntityType)
                                    {
                                        case EntityType.TextRange:
                                            paragraphText += (paraItems[i] as WTextRange).Text;
                                            break;
                                        case EntityType.Break:
                                            paragraphText += (paraItems[i] as Break).BreakType is BreakType.LineBreak ? "\r\n" : "";
                                            break;

                                    }
                                }

                                if (paraItems[i] is WFieldMark && (paraItems[i] as WFieldMark).Type is FieldMarkType.FieldEnd)
                                    isIgnoreCheckBoxItems = false;
                            }
                            //Add the question information to the list collection.
                            questionList.Add(paragraphText);
                        }
                    }
                }
                //Save the Word file.
                wordDocument.Save(Path.GetFullPath(@"../../../Output/Result.docx"));
            }
        }
    }
}
