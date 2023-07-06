using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace Replace_EMF_with_PNG_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream docStream = new FileStream(@"../../../SyncfusionConvertWordToPdfIssueDoc.docx", FileMode.Open, FileAccess.Read))
            {
                //Load file stream into Word document.
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Automatic))
                {
                    ConvertEMFToPNG(wordDocument);
                    //Instantiation of DocIORenderer for Word to PDF conversion.
                    DocIORenderer render = new DocIORenderer();
                    //Convert Word document into PDF document.
                    using (PdfDocument pdfDocument = render.ConvertToPDF(wordDocument))
                    {
                        //Save the PDF file.
                        using (FileStream outputFile = new FileStream("Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            pdfDocument.Save(outputFile);
                        
                    }
                }
            }
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.pdf")
            {
                UseShellExecute = true
            };
            process.Start();
        }

        /// <summary>
        /// Iterate through the Word document elements.
        /// </summary>
        /// <param name="wordDocument"></param>
        private static void ConvertEMFToPNG(WordDocument wordDocument)
        {
            foreach (WSection section in wordDocument.Sections)
            {
                WTextBody sectionBody = section.Body;
                //Iterate through the TextBody of the section.
                IterateTextBody(sectionBody);
                WHeadersFooters headersFooters = section.HeadersFooters;
                //Iterate through the TextBody of Header and Footer.
                IterateTextBody(headersFooters.OddHeader);
                IterateTextBody(headersFooters.OddFooter);
                IterateTextBody(headersFooters.FirstPageHeader);
                IterateTextBody(headersFooters.FirstPageFooter);
                IterateTextBody(headersFooters.EvenHeader);
                IterateTextBody(headersFooters.EvenFooter);
            }
        }
        /// <summary>
        /// Iterate through the text body elements.
        /// </summary>
        /// <param name="textBody"></param>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterate through each of the child items of WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Access the body items (should be either paragraph, table or block content control) as IEntity.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control.
                //Decide the element type by using EntityType.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Process the paragraph contents.
                        //Iterate through the paragraph's DOM.
                        IterateParagraph(paragraph.Items);
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells.
                        //Iterate through table's DOM.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterate to the body items of Block Content Control.
                        IterateTextBody(blockContentControl.TextBody);
                        break;
                }
            }
        }
        /// <summary>
        /// Iterate through the table elements.
        /// </summary>
        /// <param name="table"></param>
        private static void IterateTable(WTable table)
        {
            //Iterate the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Table cell is derived from (also a) TextBody.
                    //Reusing the code meant for iterating TextBody.
                    IterateTextBody(cell);
                }
            }
        }
        /// <summary>
        /// Iterate through the paragraph elements.
        /// </summary>
        /// <param name="paraItems"></param>
        private static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                //A paragraph can have child elements such as text, image, hyperlink, symbols, etc.,
                //Decide the element type by using EntityType.
                switch (entity.EntityType)
                {
                    case EntityType.Picture:
                        WPicture picture = entity as WPicture;
                        //Convert EMG image format to PNG format with the same size.
                        Image image = Image.FromStream(new MemoryStream(picture.ImageBytes));
                        if (image.RawFormat.Equals(ImageFormat.Emf) || image.RawFormat.Equals(ImageFormat.Wmf))
                        {
                            float height = picture.Height;
                            float width = picture.Width;
                            FileStream imgFile = new FileStream("Output.png", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                            image.Save(imgFile, ImageFormat.Png);
                            imgFile.Dispose();
                            image.Dispose();

                            FileStream imageStream = new FileStream(@"Output.png", FileMode.Open, FileAccess.ReadWrite);
                            picture.LoadImage(imageStream);
                            picture.LockAspectRatio = false;
                            picture.Height = height;
                            picture.Width = width;
                            imageStream.Dispose();
                        }
                        break;
                    case EntityType.TextBox:
                        //Iterate to the body items of textbox.
                        WTextBox textBox = entity as WTextBox;
                        IterateTextBody(textBox.TextBoxBody);
                        break;
                    case EntityType.Shape:
                        //Iterate to the body items of shape.
                        Shape shape = entity as Shape;
                        IterateTextBody(shape.TextBody);
                        break;
                    case EntityType.InlineContentControl:
                        //Iterate to the paragraph items of inline content control.
                        InlineContentControl inlineContentControl = entity as InlineContentControl;
                        IterateParagraph(inlineContentControl.ParagraphItems);
                        break;
                }
            }
        }
    }
}
