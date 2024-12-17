using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_text_in_list_paragraph
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Open the template Word document.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    string htmlFilePath = @"Data/File.html";
                    //Check if the HTML content is valid.
                    if (document.LastSection.Body.IsValidXHTML(htmlFilePath, XHTMLValidationType.None))
                    {
                        //Define the variable containing the text to search within the paragraph.
                        string variable = "Youth mountain bike";
                        //Find the first occurrence of a particular text in the document
                        TextSelection textSelection = document.Find(variable, true, true);
                        //Get the found text as single text range
                        WTextRange textRange = textSelection.GetAsOneRange();
                        // Get the paragraph containing the found text range
                        WParagraph paragraph = textRange.OwnerParagraph;
                        //Get the next sibling element of the current paragraph.
                        TextBodyItem nextSibling = paragraph.NextSibling as TextBodyItem;
                        //Get the index of the current paragraph within its parent text body.
                        int sourceIndex = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);
                        //Clear all child entities within the paragraph.
                        paragraph.ChildEntities.Clear();
                        //Get the list style name from the paragraph.
                        string listStyleName = paragraph.ListFormat.CurrentListStyle.Name;
                        //Get the current list level number.
                        int listLevelNum = paragraph.ListFormat.ListLevelNumber;
                        //Append HTML content from the specified file to the paragraph.
                        paragraph.AppendHTML(File.ReadAllText(Path.GetFullPath(htmlFilePath)));
                        //Determine the index of the next sibling if it exists.
                        int nextSiblingIndex = nextSibling != null ? nextSibling.OwnerTextBody.ChildEntities.IndexOf(nextSibling) : -1;
                        //Apply the same list style to newly added paragraphs from the HTML content.
                        for (int k = sourceIndex; k < paragraph.OwnerTextBody.Count; k++)
                        {
                            //Stop applying the style if the next sibling is reached.
                            if (nextSiblingIndex != -1 && k == nextSiblingIndex)
                            {
                                break;
                            }
                            Entity entity = paragraph.OwnerTextBody.ChildEntities[k];
                            //Apply the list style only if the entity is a paragraph.
                            if (entity is WParagraph)
                            {
                                (entity as WParagraph).ListFormat.ApplyStyle(listStyleName);
                                (entity as WParagraph).ListFormat.ListLevelNumber = listLevelNum;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the modified Word document to the output file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
