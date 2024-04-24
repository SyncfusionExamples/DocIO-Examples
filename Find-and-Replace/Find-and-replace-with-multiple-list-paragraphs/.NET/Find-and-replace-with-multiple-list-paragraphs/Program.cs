using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


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
    //Clone the paragraph.
    WParagraph paratoInsert = ownerPara.Clone() as WParagraph;
    //Creates new text body part.
    TextBodyPart bodyPart = new TextBodyPart(ownerPara.Document);
    //Iterate the body items of the sub document.
    for (int bodyItemIndex = 0; bodyItemIndex < subDocument.LastSection.Body.ChildEntities.Count; bodyItemIndex++)
    {
        if (subDocument.LastSection.Body.ChildEntities[bodyItemIndex] is WParagraph)
        {
            WParagraph paragraph = subDocument.LastSection.Body.ChildEntities[bodyItemIndex] as WParagraph;
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
    }
    return bodyPart;
}