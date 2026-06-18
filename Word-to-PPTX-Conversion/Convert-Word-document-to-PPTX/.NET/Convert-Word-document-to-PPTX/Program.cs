using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Presentation;

namespace Convert_Word_document_to_PPTX
{
    class Program
    {
        private static IPresentation pptxDoc;
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath("../../../Data/Template.docx"), Syncfusion.DocIO.FormatType.Automatic))
            {
                // Create a new PowerPoint presentation.
                pptxDoc = Presentation.Create();
                // Iterate each section in the Word document and process its body.
                foreach (WSection section in document.Sections)
                {
                    // Access the section body that contains paragraphs, tables, and content controls.
                    WTextBody sectionBody = section.Body;
                    AddTextBodyItems(sectionBody);
                }
                FileStream outputStream = new FileStream(Path.GetFullPath("../../../Output/DocxToPptx.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
                pptxDoc.Save(outputStream);
                outputStream.Close();
            }
        }

        /// <summary>
        /// Iterates the text body items of Word document and creates slides and add textbox accordingly
        /// </summary>
        /// <param name="docTextBody"></param>
        /// <param name="powerPointTableCell"></param>
        private static void AddTextBodyItems(WTextBody docTextBody, ICell powerPointTableCell = null)
        {
            ISlide powerPointSlide = null;
            IShape powerPointShape = null;
            IParagraph powerPointParagraph = null;

            //Iterates through each of the child items of WTextBody
            for (int i = 0; i < docTextBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items (should be either paragraph, table or block content control) as IEntity
                IEntity bodyItemEntity = docTextBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                //Decides the element type by using EntityType
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph docParagraph = bodyItemEntity as WParagraph;
                        if (docParagraph.ChildEntities.Count == 0)
                            break;
                        //Checkes whether the paragraph is list paragraph
                        if (IsListParagraph(docParagraph))
                        {
                            if (docParagraph.ListFormat.ListType == Syncfusion.DocIO.DLS.ListType.NoList)
                            {
                                powerPointSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
                                powerPointSlide.AddTextBox(50, 50, 756, 300);
                            }
                            powerPointShape = powerPointSlide.Shapes[0] as IShape;
                            powerPointParagraph = powerPointShape.TextBody.AddParagraph(docParagraph.Text);
                            //Checks whether the list type is numbered
                            if (docParagraph.ListFormat.ListType == Syncfusion.DocIO.DLS.ListType.Numbered)
                                ApplyListStyles(powerPointParagraph);
                        }
                        //Checks whether the paragraph is inside a cell
                        else if (docParagraph.IsInCell && powerPointTableCell != null)
                        {
                            powerPointParagraph = powerPointTableCell.TextBody.AddParagraph();
                            AddParagraphItems(docParagraph, powerPointParagraph, powerPointTableCell);
                        }
                        //Checks whether the paragraph is heading style
                        else if (docParagraph.StyleName.Contains("Heading"))
                        {
                            powerPointSlide = pptxDoc.Slides.Add(SlideLayoutType.Title);
                            powerPointSlide.Shapes.Remove(powerPointSlide.Shapes[1]);
                            powerPointShape = powerPointSlide.Shapes[0] as IShape;
                            powerPointParagraph = powerPointShape.TextBody.AddParagraph();

                            AddParagraphItems(docParagraph, powerPointParagraph);
                        }
                        else
                        {
                            powerPointSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
                            powerPointShape = powerPointSlide.AddTextBox(50, 50, 800, 300);
                            powerPointParagraph = powerPointShape.TextBody.AddParagraph();
                            powerPointParagraph.FirstLineIndent = 30;

                            AddParagraphItems(docParagraph, powerPointParagraph);
                        }
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells
                        //Iterates through table's DOM
                        //Iterates the row collection in a table and creates a new slide for each row
                        WTable table = bodyItemEntity as WTable;
                        foreach (WTableRow row in table.Rows)
                        {
                            powerPointSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
                            ITable powerPointTable = powerPointSlide.Shapes.AddTable(1, row.Cells.Count, 200, 140, 500, 220);
                            powerPointTable.BuiltInStyle = BuiltInTableStyle.None;
                            //Iterates the cell collection in a table row
                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                WTableCell cell = row.Cells[j];
                                //Table cell is derived from (also a) TextBody
                                //Reusing the code meant for iterating TextBody
                                AddTextBodyItems(cell as WTextBody, powerPointTable.Rows[0].Cells[j]);
                            }
                        }
                        break;
                    case EntityType.BlockContentControl:
                        BlockContentControl blockContentControl = bodyItemEntity as BlockContentControl;
                        //Iterates to the body items of Block Content Control.
                        AddTextBodyItems(blockContentControl.TextBody);
                        break;
                }
            }
        }

        /// <summary>
        /// Applies the Numbered list style 
        /// </summary>
        /// <param name="powerPointParagraph"></param>
        private static void ApplyListStyles(IParagraph powerPointParagraph)
        {
            powerPointParagraph.ListFormat.Type = Syncfusion.Presentation.ListType.Numbered;
            powerPointParagraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
            powerPointParagraph.ListFormat.StartValue = 1;
            // Sets the hanging value
            powerPointParagraph.FirstLineIndent = -20;
            // Sets the bullet character size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
            powerPointParagraph.ListFormat.Size = 100;
            powerPointParagraph.IndentLevelNumber = 1;
        }

        /// <summary>
        /// Iterates the paragraph and adds the paragraph items to the Presentation document
        /// </summary>
        /// <param name="docParagraph"></param>
        /// <param name="powerPointParagraph"></param>
        /// <param name="powerPointTableCell"></param>
        private static void AddParagraphItems(WParagraph docParagraph, IParagraph powerPointParagraph, ICell powerPointTableCell = null)
        {
            for (int i = 0; i < docParagraph.Items.Count; i++)
            {
                Entity entity = docParagraph.Items[i];
                //A paragraph can have child elements such as text, image, hyperlink, symbols, etc.,
                //Decides the element type by using EntityType
                switch (entity.EntityType)
                {
                    case EntityType.TextRange:
                        WTextRange textRange = entity as WTextRange;
                        ITextPart textPart = powerPointParagraph.AddTextPart(textRange.Text);
                        //Checks whether th paragraph is not in cell
                        if (!docParagraph.IsInCell)
                        {
                            //Checks whether the paragraph is heading style paragraph
                            if (docParagraph.StyleName.Contains("Heading"))
                                textPart.Font.Bold = true;
                            else
                            {
                                textPart.Font.Color.SystemColor = Color.Black;
                                textPart.Font.FontSize = 32;
                            }
                        }
                        break;
                    case EntityType.Picture:
                        //Checks whether the image is inside a cell
                        if (docParagraph.IsInCell && powerPointTableCell != null)
                            powerPointTableCell.Fill.PictureFill.ImageBytes = (entity as WPicture).ImageBytes;
                        break;
                }
            }
        }

        /// <summary>
        /// Checks whether the paragraph is list paragraph
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        private static bool IsListParagraph(WParagraph paragraph)
        {
            return paragraph.NextSibling is WParagraph && paragraph.PreviousSibling is WParagraph &&
                ((paragraph.NextSibling as WParagraph).ListFormat.ListType != Syncfusion.DocIO.DLS.ListType.NoList
                || (paragraph.PreviousSibling as WParagraph).ListFormat.ListType != Syncfusion.DocIO.DLS.ListType.NoList);
        }
    }
}