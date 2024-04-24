using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Reflection.Metadata;


string[] filePaths = {"../../../Data/Heading1Items.docx","../../../Data/Heading2Items.docx"};
//Open the file as Stream.
using (FileStream documentStream = new FileStream("../../../Data/Input.docx", FileMode.Open, FileAccess.Read))
{
    //Open an existing Word document.
    using (WordDocument document = new WordDocument(documentStream, FormatType.Docx))
    {
        for(int documentIndex = 1; documentIndex <= filePaths.Length; documentIndex++)
        {
            //String to be found.
            string findText = "<<Heading"+ documentIndex + "Items>>";
            //Find the selection in the document.
            TextSelection selection = document.Find(findText, false, false);
            //Get the owner paragraph.
            WParagraph ownerPara = selection.GetAsOneRange().OwnerParagraph;
            //Open the file as Stream.
            using (FileStream subDocumentStream = new FileStream(filePaths[documentIndex-1], FileMode.Open, FileAccess.Read))
            {
                //Open an sub Word document.
                using (WordDocument subDocument = new WordDocument(subDocumentStream, FormatType.Docx))
                {
                    //Create a text body part to be replaced.
                    TextBodyPart textBodyPart = CreateBodyPart(subDocument, ownerPara);
                    //Replace the text with the created text body part.
                    document.Replace(findText, textBodyPart, true, true);
                }
            }
        }

        //Save the Word document.
        using (FileStream output = new FileStream("../../../Result.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(output, FormatType.Docx);
        }
    }
}

/// <summary>
/// Create body parts that need to be replaced in the Word document.
/// </summary>
/// <param name="subDocument">Document contains the paragraph which need to be replace</param>
/// <param name="ownerPara">The paragraph that needs to be copied.</param>
TextBodyPart CreateBodyPart(WordDocument subDocument, WParagraph ownerPara)
{
    //Creates new text body part.
    TextBodyPart bodyPart = new TextBodyPart(ownerPara.Document);
    //Iterate each section of the Word document.
    foreach (WSection section in subDocument.Sections)
    {
        //Accesses the Body of section where all the contents in document are apart
        WTextBody sectionBody = section.Body;
        IterateTextBody(sectionBody, ownerPara, bodyPart);
    }
    return bodyPart;
}
/// <summary>
/// Iterates textbody child elements.
/// </summary>
void IterateTextBody(WTextBody sectionTextBody, WParagraph ownerPara, TextBodyPart bodyPart)
{
    //Iterates through each of the child items of WTextBody
    for (int i = 0; i < sectionTextBody.ChildEntities.Count; i++)
    {
        //IEntity is the basic unit in DocIO DOM. 
        //Accesses the body items (should be either paragraph, table or block content control) as IEntity
        IEntity bodyItemEntity = sectionTextBody.ChildEntities[i];
        //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
        //Decides the element type by using EntityType
        switch (bodyItemEntity.EntityType)
        {
            case EntityType.Paragraph:
                WParagraph paragraph = bodyItemEntity as WParagraph;
                AddParagraphItemsToTextBody (paragraph, ownerPara, bodyPart);
                break;
            case EntityType.Table:
                //Add to text body part.
                bodyPart.BodyItems.Add(bodyItemEntity.Clone());
                break;
            case EntityType.BlockContentControl:
                BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                //Iterates to the body items of Block Content Control.
                IterateTextBody(blockContentControl.TextBody, ownerPara, bodyPart);
                break;
        }
    }
}
/// <summary>
/// Add the paragraph items to the text body part.
/// </summary>
void AddParagraphItemsToTextBody(WParagraph paragraph, WParagraph ownerPara, TextBodyPart bodyPart)
{
    //Clone the paragraph.
    WParagraph paratoInsert = ownerPara.Clone() as WParagraph;
    //Clear the paragraph's items.
    paratoInsert.ChildEntities.Clear();
    //Iterate the items of the paragraph.
    for (int paraItemIndex = 0; paraItemIndex < paragraph.ChildEntities.Count; paraItemIndex++)
    {
        //Add the paragraph items.
        paratoInsert.ChildEntities.Add(paragraph.ChildEntities[paraItemIndex].Clone());
    }
    //Add to text body part.
    bodyPart.BodyItems.Add(paratoInsert.Clone());
}