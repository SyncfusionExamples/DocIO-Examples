using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_list_text_with_paragraphs
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream destDocStream = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the destination Word document.
                using (WordDocument destDocument = new WordDocument(destDocStream, FormatType.Docx))
                {
                    using (FileStream sourceDocStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source Word document.
                        using (WordDocument sourceDocument = new WordDocument(sourceDocStream, FormatType.Docx))
                        {
                            //Find the text "[finder]" in the destination document.
                            TextSelection textSelection = destDocument.Find("[finder]", false, false);
                            if (textSelection != null)
                            {
                                //Get the paragraph containing the found text in destintaion document.
                                WParagraph destOwnerParagraph = textSelection.GetAsOneRange().OwnerParagraph;
                                //Get the owner section.
                                WSection destOwnerSection = destOwnerParagraph.OwnerTextBody.Owner as WSection;
                                //Get the owner paragraph index.
                                int destOwnerParaIndex = destOwnerSection.Paragraphs.IndexOf(destOwnerParagraph);
                                //Retrieve the first section of the source document.
                                WSection section = sourceDocument.Sections[0];
                                //Iterate through the each paragraph of the source document.
                                for (int i = 0; i < section.Paragraphs.Count; i++)
                                {
                                    WParagraph sourcePara = section.Paragraphs[i];
                                    //Replace the found text with the first paragraph text from the source document.
                                    if (i == 0)
                                        destOwnerParagraph.Replace(textSelection.SelectedText, sourcePara.Text, false, false);
                                    //For remaining paragraphs in the source document.
                                    else
                                    {
                                        //Clone the found destination paragraph.
                                        WParagraph paraToInsert = (WParagraph)destOwnerParagraph.Clone();
                                        //Change the paragraph text to the source paragraph text.
                                        paraToInsert.Text = sourcePara.Text;
                                        //Insert the cloned paragraph as the next item of the found destination paragraph.
                                        destOwnerSection.Body.ChildEntities.Insert(destOwnerParaIndex, paraToInsert);
                                    }
                                    //Increase the paragraph index to insert the next cloned paragraph in the correct position.
                                    destOwnerParaIndex++;
                                }
                            }
                            //Save the destination Word document.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Saves the Word document to file stream.
                                destDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
