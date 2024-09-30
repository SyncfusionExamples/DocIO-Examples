using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


//Open the existing main document.
FileStream fileStream1 = new FileStream(Path.GetFullPath(@"Data/Main.docx"), FileMode.Open);
WordDocument mainDocument = new WordDocument(fileStream1, FormatType.Docx);
//Open the existing temporary document.
FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/REKOMENDASI.docx"), FileMode.Open);
WordDocument tempDoc = new WordDocument(fileStream, FormatType.Docx);
//Set the first section break.
tempDoc.Sections[0].BreakCode = SectionBreakCode.NoBreak;

//Get the last section index of main document.
int secIndex = mainDocument.ChildEntities.IndexOf(mainDocument.LastSection);
//Get the last paragraph index of main document.
int paraIndex = mainDocument.LastSection.Body.ChildEntities.IndexOf(mainDocument.LastParagraph);
//Get the last paragraph style.
WParagraph lastPara = mainDocument.LastParagraph;

//Import the temporary document content to the main document.
mainDocument.ImportContent(tempDoc, ImportOptions.UseDestinationStyles);

//Modify the paragraph style for the added contents.
AddLeftIndentation(mainDocument, secIndex, paraIndex + 1, lastPara.ParagraphFormat.LeftIndent);

//Save the main document.
FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output_REKOMENDASI.docx"), FileMode.Create, FileAccess.Write);
mainDocument.Save(outputStream, FormatType.Docx);


void AddLeftIndentation(WordDocument document1, int secIndex, int paraIndex, float leftIndent)
{
    //Iterate through the sections added from the temporary document.
    for (int i = secIndex; i < document1.ChildEntities.IndexOf(document1.LastSection) + 1; i++)
    {
        //Iterate through the child entities added from the temporary document.
        for (int j = paraIndex; j < document1.Sections[i].Body.ChildEntities.Count; j++)
        {
            //If the child entity is a paragraph then apply the previously taken para left indent.
            if (document1.Sections[i].Body.ChildEntities[j] is WParagraph)
            {
                WParagraph para = document1.Sections[i].Body.ChildEntities[j] as WParagraph;
                //Set the left indentation
                para.ParagraphFormat.LeftIndent = leftIndent;
            }
            //If the child entity is a table then apply the previously taken para left indent.
            else if (document1.Sections[i].Body.ChildEntities[j] is WTable)
            {
                WTable table = document1.Sections[i].Body.ChildEntities[j] as WTable;
                //Set the left indentation
                table.TableFormat.LeftIndent = leftIndent;
            }
        }
        //Reset the index for next section.
        paraIndex = 0;
    }
}