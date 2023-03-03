using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Find_all_items_by_properties
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    string[] propertyNames = { "ContentControlProperties.Title", "ContentControlProperties.Tag" };
                    string[] propertyValues = { "CompanyName", "CompanyName" };

                    //Find all block content controls by Title and Tag. 
                    List<Entity> blockContentControls = document.FindAllItemsByProperties(EntityType.BlockContentControl, propertyNames, propertyValues);

                    //Iterates the block content controls and remove the block content controls.
                    for (int i = 0; i < blockContentControls.Count; i++)
                    {
                        BlockContentControl blockContentControl = blockContentControls[i] as BlockContentControl;
                        blockContentControl.OwnerTextBody.ChildEntities.Remove(blockContentControl);
                    }
                    propertyNames = new string[] { "ContentControlProperties.Title", "ContentControlProperties.Tag" };
                    propertyValues = new string[] { "Contact", "Contact" };

                    //Find all inline content controls by Title and Tag. 
                    List<Entity> inlineContentControls = document.FindAllItemsByProperties(EntityType.InlineContentControl, propertyNames, propertyValues);

                    //Iterates the inline content controls and remove the inline content controls.
                    for (int i = 0; i < inlineContentControls.Count; i++)
                    {
                        InlineContentControl inlineContentControl = inlineContentControls[i] as InlineContentControl;
                        inlineContentControl.OwnerParagraph.ChildEntities.Remove(inlineContentControl);
                    }
                    propertyNames = new string[] { "CharacterFormat.Bold", "CharacterFormat.Italic" };
                    propertyValues = new string[] { true.ToString(), true.ToString() };

                    //Find all bold and italic text.
                    List<Entity> textRanges = document.FindAllItemsByProperties(EntityType.TextRange, propertyNames, propertyValues);

                    //Iterates the textRanges and remove the bold and italic.
                    for (int i = 0; i < textRanges.Count; i++)
                    {
                        WTextRange textRange = textRanges[i] as WTextRange;
                        textRange.CharacterFormat.Bold = false;
                        textRange.CharacterFormat.Italic = false;
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
