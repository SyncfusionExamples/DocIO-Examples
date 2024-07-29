using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(@"../../../Data/InsertComment.docx", FileMode.Open, FileAccess.Read))
{
    //Open the existing Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        //Find all occurrence of a particular text ending with comma in the document using regex.
        TextSelection[] textSelection = document.FindAll("panda", true, true);
        if (textSelection != null)
        {
            //Iterates through each occurrence and comment it.
            for (int i = 0; i < textSelection.Count(); i++)
            {
                WTextRange textRange = textSelection[i].GetAsOneRange();

                //Get the index of the found text.
                int textIndex = textRange.OwnerParagraph.ChildEntities.IndexOf(textRange);
                //Add comment to a paragraph.
                WComment comment = textRange.OwnerParagraph.AppendComment("comment test_" + i);
                //Specify the author of the comment.
                comment.Format.User = "Peter";
                //Set the date and time for the comment.
                comment.Format.DateTime = DateTime.Now;
                //Insert the comment next to the textrange.
                textRange.OwnerParagraph.ChildEntities.Insert(textIndex + 1, comment);
                //Add the paragraph items to the commented items.
                comment.AddCommentedItem(textRange);
            }
        }

        //Save the Word document
        using (FileStream outputStream = new FileStream(@"../../../Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}