using Syncfusion.DocIO.DLS;

// Load the Word document using Syncfusion.DocIO
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Define an array of texts to find within the document
    string[] textsToFind = new string[2] { "GIANT START", "GIANT END" };

    // Iterate through each text in the array and call FindStartEndAndIterate to process them
    foreach (string textToFind in textsToFind)
    {
        FindStartEndAndIterate(document, textToFind);
    }
}

/// <summary>
/// Finds the start and end paragraphs based on the search text, then iterates through the paragraphs in the specified range.
/// </summary>
/// <param name="document">The Word document to search through.</param>
/// <param name="textToSearch">The text to search for that marks the start of the range.</param>
void FindStartEndAndIterate(WordDocument document, string textToSearch)
{
    WParagraph startPara = null;
    WParagraph endPara = null;
    WSection startSection = null;
    WSection endSection = null;

    string endText = "GIANT"; // Text to mark the end of the range

    // Step 1: Find the start and end paragraphs by iterating through each section of the document
    foreach (WSection section in document.Sections)
    {
        foreach (Entity entity in section.Body.ChildEntities)
        {
            // Check if the entity is a paragraph
            if (entity is WParagraph paragraph)
            {
                string paraText = paragraph.Text;

                // If the paragraph starts with the 'textToSearch', set it as the start paragraph
                if (paraText.StartsWith(textToSearch))
                {
                    startPara = paragraph;
                    startSection = section;
                }
                // If the paragraph contains the 'endText', set it as the end paragraph
                else if (paraText.Contains(endText))
                {
                    endPara = paragraph;
                    endSection = section;
                    break; // Stop once the end paragraph is found
                }
            }
        }

        // If both start and end paragraphs have been found, break out of the loop
        if (startPara != null && endPara != null)
            break;
    }

    // If no start or end paragraphs were found, exit the method
    if (startPara == null || endPara == null)
        return;

    // Get the index of the start paragraph within its section's body and the section itself
    int startBodyIndex = startSection.Body.ChildEntities.IndexOf(startPara);
    int startSectionIndex = document.Sections.IndexOf(startSection);

    // Get the index of the end paragraph within its section's body and the section itself
    int endBodyIndex = endSection.Body.ChildEntities.IndexOf(endPara);
    int endSectionIndex = document.Sections.IndexOf(endSection);

    // Step 2: Loop through the sections from the start section to the end section
    for (int sectionIndex = startSectionIndex; sectionIndex <= endSectionIndex; sectionIndex++)
    {
        // Get the current section from the document
        WSection currentSection = document.Sections[sectionIndex];

        // Determine the starting index of body entities in this section
        // If it's the first section, start from startBodyIndex; otherwise, start from the beginning
        int start = (sectionIndex == startSectionIndex) ? startBodyIndex : 0;

        // Determine the ending index of body entities in this section
        // If it's the last section, end at endBodyIndex; otherwise, end at the last entity
        int end = (sectionIndex == endSectionIndex) ? endBodyIndex : currentSection.Body.ChildEntities.Count - 1;

        // Step 3: Loop through the paragraphs from start to end within the current section
        for (int paraIndex = start; paraIndex <= end; paraIndex++)
        {
            // Get the current entity in the section's body
            Entity currEntity = currentSection.Body.ChildEntities[paraIndex];

            // Check if the current entity is a paragraph
            if (currEntity is WParagraph currentPara)
            {
                // Print the paragraph text to the console
                Console.WriteLine(currentPara.Text);
            }
        }
    }
    Console.ReadLine();
}

