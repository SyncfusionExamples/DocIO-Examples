using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_line_break_with_paragraph_mark
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    // Iterate through the paragraphs of Word document.
                    foreach (WParagraph paragraph in document.Sections[0].Body.Paragraphs)
                    {
                        for (int i = 0; i < paragraph.Items.Count; i++)
                        {
                            Entity entity = paragraph.Items[i];
                            //Set paragraph child element as break.                                                     
                            if (entity.EntityType == EntityType.Break)
                            {
                                Break breakItem = entity as Break;
                                //Replace line break with paragraph mark.
                                if (breakItem.BreakType == BreakType.LineBreak)
                                {

                                    WParagraph ownerPara = breakItem.OwnerParagraph;
                                    int breakIndex = ownerPara.ChildEntities.IndexOf(breakItem);
                                    int paraIndex = ownerPara.OwnerTextBody.ChildEntities.IndexOf(ownerPara);

                                    //Create new paragraph by cloning the existing paragraph.
                                    WParagraph newPara = ownerPara.Clone() as WParagraph;

                                    //Remove the child items after the line break from the old paragraph including line break.
                                    for (int j = breakIndex; j < ownerPara.ChildEntities.Count;)
                                    {
                                        ownerPara.ChildEntities.RemoveAt(j);
                                    }
                                    int newParaItemsCount = ownerPara.ChildEntities.Count;
                                    //Remove the child items before the line break from the new paragraph including line break.
                                    while (newParaItemsCount + 1 != 0)
                                    {
                                        newPara.ChildEntities.RemoveAt(0);
                                        newParaItemsCount--;
                                    }
                                    //Insert the new paragraph next to the line break paragraph.
                                    ownerPara.OwnerTextBody.ChildEntities.Insert(paraIndex + 1, newPara);

                                }
                            }
                        }
                    }
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
