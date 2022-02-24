using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Split_by_heading
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read))
            {
                //Load the template document as stream
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    inputStream.Dispose();
                    WordDocument newDocument = null;
                    WSection newSection = null;
                    int i = 0;
                    //Iterate each section from Word document
                    foreach (WSection section in document.Sections)
                    {
                        foreach (TextBodyItem textbodyitem in section.Body.ChildEntities)
                        {
                            if (textbodyitem is WParagraph)
                            {
                                WParagraph para = textbodyitem as WParagraph;
                                if (para.StyleName == "Heading 1")
                                {
                                    if (newDocument != null)
                                    {
                                        //Saves the Word document to  MemoryStream
                                        using (FileStream outputStream = new FileStream(@"../../../Heading" + i + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                                        {
                                            newDocument.Save(outputStream, FormatType.Docx);
                                            //Closes the document
                                            newDocument.Close();
                                            newDocument = null;
                                        }
                                        i++;
                                    }
                                    //Create new Word document
                                    newDocument = new WordDocument();
                                    newSection = newDocument.AddSection() as WSection;
                                    //Add cloned paragraphs into new section
                                    newSection.Body.ChildEntities.Add(para.Clone());
                                }
                                else if (para.StyleName != "Heading 1" && newDocument != null)
                                {
                                    //Add cloned paragraphs into new section
                                    newSection.Body.ChildEntities.Add(para.Clone());
                                }
                            }
                            else if (textbodyitem is WTable)
                            {
                                //Add cloned table into new section
                                WTable table = textbodyitem as WTable;
                                newSection.Body.ChildEntities.Add(table.Clone());
                            }
                            else if (textbodyitem is BlockContentControl)
                            {
                                //Add cloned block content control into new section
                                BlockContentControl contentControl = textbodyitem as BlockContentControl;
                                newSection.Body.ChildEntities.Add(contentControl.Clone());
                            }
                        }
                    }
                    if (newDocument != null)
                    {
                        //Saves the Word document to  MemoryStream
                        using (FileStream outputStream = new FileStream(@"../../../Heading" + i + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                        {
                            newDocument.Save(outputStream, FormatType.Docx);
                            //Closes the document
                            newDocument.Close();
                        }
                    }
                }
            }
        }
    }
}
