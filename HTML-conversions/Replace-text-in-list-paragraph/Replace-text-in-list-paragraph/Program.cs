using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Register Syncfusion license
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Mgo+DSMBMAY9C3t2UlhhQlNHfV5DQmBWfFN0QXNYfVRwdF9GYEwgOX1dQl9nSXZTc0VlWndfcXNSQWc=");

using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Open the template Word document.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
    {
        string htmlFilePath = @"Data/File.html";
        //Check if the HTML content is valid.
        bool isvalidHTML = IsValidHTML(htmlFilePath);
        if (isvalidHTML)
        {
            //Iterate through the sections in the document.
            for (int i = 0; i < document.Sections.Count; i++)
            {
                //Iterate through the paragraphs within the section.
                for (int j = 0; j < document.Sections[i].Paragraphs.Count; j++)
                {
                    //Get the current paragraph from the section.
                    WParagraph paragraph = document.Sections[i].Paragraphs[j] as WParagraph;
                    //Define the variable containing the text to search within the paragraph.
                    string variable = "Mountain-300";
                    //If the paragraph contains the specific text, replace it with HTML content.
                    if (paragraph.Text.Contains(variable))
                    {
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
                        //Reapply the original list style to the paragraph.
                        paragraph.ListFormat.ApplyStyle(listStyleName);
                        //Reapply the original list level number.
                        paragraph.ListFormat.ListLevelNumber = listLevelNum;
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
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
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

/// <summary>
/// Validates whether the HTML content from the specified file is well-formed XHTML.
/// </summary>
static bool IsValidHTML(string htmlFilePath)
{
    using (WordDocument document = new WordDocument())
    {
        //Add a section for the HTML content.
        IWSection section = document.AddSection();
        //Read the HTML string from the specified file.
        string htmlString = File.ReadAllText(Path.GetFullPath(htmlFilePath));
        //Validate the HTML string.
        return section.Body.IsValidXHTML(htmlString, XHTMLValidationType.None);
    }
}
