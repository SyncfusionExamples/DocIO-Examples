using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.Formatting;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Format_word_document_uniformly
{
    class Program
    {
        // Holds all formatting standards that will be applied
        // throughout the document.
        static FormattingRules rules = new FormattingRules
        {
        };

        static void Main(string[] args)
        {
            // Load the input Word document.
            using (WordDocument document = new WordDocument(
                Path.GetFullPath("../../../Data/Template.docx"),
                FormatType.Automatic))
            {
                // Iterate through all sections in the document.
                foreach (WSection section in document.Sections)
                {
                    // Process document body content.
                    IterateTextBody(section.Body);

                    // Process headers and footers.
                    WHeadersFooters headersFooters = section.HeadersFooters;

                    IterateTextBody(headersFooters.OddHeader);
                    IterateTextBody(headersFooters.OddFooter);
                    IterateTextBody(headersFooters.EvenHeader);
                    IterateTextBody(headersFooters.EvenFooter);
                    IterateTextBody(headersFooters.FirstPageHeader);
                    IterateTextBody(headersFooters.FirstPageFooter);
                }

                // Save the formatted document.
                using (FileStream outputFileStream = new FileStream(
                    Path.GetFullPath(@"../../../Output/output.docx"),
                    FileMode.Create,
                    FileAccess.ReadWrite))
                {
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        /// <summary>
        /// Iterates through all entities present inside a text body
        /// such as paragraphs, tables, and content controls.
        /// </summary>
        static void IterateTextBody(WTextBody textBody)
        {
            if (textBody == null)
                return;

            Stack<int> HeadingStack = new Stack<int>();

            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                IEntity bodyItemEntity = textBody.ChildEntities[i];

                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        {
                            WParagraph paragraph = bodyItemEntity as WParagraph;

                            // Validate and correct heading hierarchy.
                            if (paragraph.StyleName.StartsWith("Heading"))
                            {
                                string styleName = paragraph.StyleName?.Trim();

                                if (!string.IsNullOrEmpty(styleName) &&
                                    styleName.StartsWith("Heading ") &&
                                    int.TryParse(styleName.Replace("Heading ", ""), out int styleLevel))
                                {
                                    int correctLevel = styleLevel;

                                    if (HeadingStack.Count == 0)
                                    {
                                        // First heading encountered.
                                        HeadingStack.Push(correctLevel);
                                    }
                                    else
                                    {
                                        int previousLevel = HeadingStack.Peek();

                                        // Example: Heading 1 -> Heading 3
                                        // Correct to Heading 2.
                                        if (styleLevel > previousLevel + 1)
                                        {
                                            correctLevel = previousLevel + 1;
                                            HeadingStack.Push(correctLevel);
                                        }
                                        // Same heading level.
                                        else if (styleLevel == previousLevel)
                                        {
                                            // No action required.
                                        }
                                        // Valid level progression.
                                        else if (styleLevel > previousLevel)
                                        {
                                            HeadingStack.Push(styleLevel);
                                        }
                                        // Moving back to a previous heading level.
                                        else
                                        {
                                            while (HeadingStack.Count > 0 &&
                                                   HeadingStack.Peek() > styleLevel)
                                            {
                                                HeadingStack.Pop();
                                            }
                                        }
                                    }

                                    // Apply corrected heading style if required.
                                    if (correctLevel != styleLevel)
                                    {
                                        paragraph.ApplyStyle($"Heading {correctLevel}");
                                    }
                                }
                            }
                            else
                            {
                                // Standardize paragraph spacing.
                                if (paragraph.ParagraphFormat.AfterSpacing != rules.DefaultParagraphAfterSpacing)
                                    paragraph.ParagraphFormat.AfterSpacing = rules.DefaultParagraphAfterSpacing;

                                if (paragraph.ParagraphFormat.BeforeSpacing != rules.DefaultParagraphBeforeSpacing)
                                    paragraph.ParagraphFormat.BeforeSpacing = rules.DefaultParagraphBeforeSpacing;

                                // Standardize paragraph alignment.
                                if (paragraph.ParagraphFormat.HorizontalAlignment != rules.DefaultParaAlignment)
                                    paragraph.ParagraphFormat.HorizontalAlignment = rules.DefaultParaAlignment;

                                // Standardize list formatting.
                                if (paragraph.ListFormat.CurrentListLevel != null)
                                {
                                    paragraph.BreakCharacterFormat.FontName = rules.DefaultFontFamily;
                                    paragraph.BreakCharacterFormat.FontSize = rules.DefaultFontSize;
                                    paragraph.BreakCharacterFormat.FontSizeBidi = rules.DefaultFontSize;
                                }

                                // Process all paragraph items.
                                IterateParagraph(paragraph.Items);
                            }

                            break;
                        }

                    case EntityType.Table:
                        {
                            WTable table = bodyItemEntity as WTable;

                            // Standardize table formatting.
                            table.TableFormat.WrapTextAround = rules.DefaultTableTextWrap;
                            table.TableFormat.LeftIndent = rules.DefaultTableLeftIndent;
                            table.TableFormat.Paddings.Top = rules.DefaultTablePaddingTop;
                            table.TableFormat.Paddings.Bottom = rules.DefaultTablePaddingBottom;
                            table.TableFormat.Paddings.Left = rules.DefaultTablePaddingLeft;
                            table.TableFormat.Paddings.Right = rules.DefaultTablePaddingRight;

                            IterateTable(table);
                            break;
                        }

                    case EntityType.BlockContentControl:
                        {
                            // Process content inside block content controls.
                            BlockContentControl blockContentControl =
                                bodyItemEntity as BlockContentControl;

                            IterateTextBody(blockContentControl.TextBody);
                            break;
                        }
                }
            }
        }

        /// <summary>
        /// Iterates through tables and standardizes
        /// cell, row, and paragraph formatting.
        /// </summary>
        static void IterateTable(WTable table)
        {
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    // Standardize cell padding.
                    cell.CellFormat.Paddings.Top = rules.DefaultCellPaddingTop;
                    cell.CellFormat.Paddings.Bottom = rules.DefaultCellPaddingBottom;
                    cell.CellFormat.Paddings.Left = rules.DefaultCellPaddingLeft;
                    cell.CellFormat.Paddings.Right = rules.DefaultCellPaddingRight;

                    for (int p = 0; p < cell.ChildEntities.Count; p++)
                    {
                        IEntity entity = cell.ChildEntities[p];

                        if (entity.EntityType == EntityType.Paragraph)
                        {
                            WParagraph cellParagraph = (WParagraph)entity;

                            // Standardize paragraph spacing inside table cells.
                            if (cellParagraph.ParagraphFormat.BeforeSpacing != rules.DefaultCellParagraphSpacingBefore)
                                cellParagraph.ParagraphFormat.BeforeSpacing = rules.DefaultCellParagraphSpacingBefore;

                            if (cellParagraph.ParagraphFormat.AfterSpacing != rules.DefaultCellParagraphSpacingAfter)
                                cellParagraph.ParagraphFormat.AfterSpacing = rules.DefaultCellParagraphSpacingAfter;

                            // Standardize alignment inside cells.
                            if (cellParagraph.ParagraphFormat.HorizontalAlignment != rules.DefaultCellTextAlignment)
                                cellParagraph.ParagraphFormat.HorizontalAlignment = rules.DefaultCellTextAlignment;

                            IterateParagraph(cellParagraph.Items);
                        }
                        else
                        {
                            // Process nested content.
                            IterateTextBody(cell);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Iterates through paragraph items and standardizes
        /// fonts, text boxes, shapes, and pictures.
        /// </summary>
        static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];

                switch (entity.EntityType)
                {
                    case EntityType.Field:
                        {
                            WField field = entity as WField;

                            // Standardize field font.
                            if (field.CharacterFormat.FontSize != rules.DefaultFontSize)
                                field.CharacterFormat.FontSize = rules.DefaultFontSize;

                            if (field.CharacterFormat.FontName != rules.DefaultFontFamily)
                                field.CharacterFormat.FontName = rules.DefaultFontFamily;

                            break;
                        }

                    case EntityType.TextRange:
                        {
                            WTextRange textRange = entity as WTextRange;

                            // Standardize text font.
                            if (textRange.CharacterFormat.FontSize != rules.DefaultFontSize)
                                textRange.CharacterFormat.FontSize = rules.DefaultFontSize;

                            if (textRange.CharacterFormat.FontName != rules.DefaultFontFamily)
                                textRange.CharacterFormat.FontName = rules.DefaultFontFamily;

                            break;
                        }

                    case EntityType.TextBox:
                        {
                            // Process textbox contents recursively.
                            WTextBox textBox = entity as WTextBox;
                            IterateTextBody(textBox.TextBoxBody);
                            break;
                        }

                    case EntityType.Shape:
                        {
                            // Process shape text contents recursively.
                            Shape shape = entity as Shape;
                            IterateTextBody(shape.TextBody);
                            break;
                        }

                    case EntityType.InlineContentControl:
                        {
                            // Process inline content control contents.
                            InlineContentControl inlineContentControl =
                                entity as InlineContentControl;

                            IterateParagraph(inlineContentControl.ParagraphItems);
                            break;
                        }

                    case EntityType.Picture:
                        {
                            // Standardize image dimensions.
                            WPicture picture = entity as WPicture;

                            picture.Height = rules.DefaultImageHeight;
                            picture.Width = rules.DefaultImageWidth;

                            break;
                        }
                }
            }
        }
    }
}