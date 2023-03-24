using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Find_all_items_by_property
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
                    //Find all footnote and endnote by EntityType in Word document.
                    List<Entity> footNotes = document.FindAllItemsByProperty(EntityType.Footnote, null, null);
                    //Remove the footnotes and endnotes.
                    for (int i = 0; i < footNotes.Count; i++)
                    {
                        WFootnote footnote = footNotes[i] as WFootnote;
                        footnote.OwnerParagraph.ChildEntities.Remove(footnote);
                    }

                    //Find all fields by FieldType.
                    List<Entity> fields = document.FindAllItemsByProperty(EntityType.Field, "FieldType", FieldType.FieldHyperlink.ToString());
                    //Iterate the hyperlink field and change URL.
                    for (int i = 0; i < fields.Count; i++)
                    {
                        //Creates hyperlink instance from field to manipulate the hyperlink.
                        Hyperlink hyperlink = new Hyperlink(fields[i] as WField);
                        //Modifies the Uri of the hyperlink.
                        if (hyperlink.Type == HyperlinkType.WebLink && hyperlink.TextToDisplay == "HTML")
                            hyperlink.Uri = "http://www.w3schools.com/";
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
