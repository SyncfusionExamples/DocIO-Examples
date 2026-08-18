using Syncfusion.Drawing;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Format_Page_Number_In_Word_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            // Generates a Word document with multiple sections and page numbering.
            using (WordDocument document = GenerateLetterWithHeader())
            {
                // Saves the Word document to the output folder.
                document.Save(Path.GetFullPath(@"Output/Sample.docx"), FormatType.Docx);
            }
        }
        /// <summary>
        /// Generates a Word document with multiple sections and page numbering.
        /// </summary>
        /// <returns>Return WordDocument</returns>
        public static WordDocument GenerateLetterWithHeader()
        {
            // Sample content for the first section.
            string letterSection1 = "<p>Your cooperation in this effort is greatly appreciated</p>";

            // Sample content for the second section.
            string letterSection2 = "<p>Your cooperation in this effort is greatly appreciated</p>";

            // Loads the template Word document.
            string dataPath = Path.GetFullPath(@"Data/Template.docx");
            WordDocument letterDoc = new WordDocument(dataPath);

            // Retrieves the Normal style and updates its font settings.
            WParagraphStyle style = letterDoc.Styles.FindByName("Normal") as WParagraphStyle;

            if (style != null)
            {
                style.CharacterFormat.FontName = "Verdana";
                style.CharacterFormat.FontSize = 12;
            }

            // Defines the left margin value for the document.
            float letterMarginLeft = 140;

            // Disables XHTML validation when inserting XHTML content.
            letterDoc.XHTMLValidateOption = XHTMLValidationType.None;

            // Tracks the current section number.
            int sectionNumber = 1;

            // Adds a new section to the document.
            letterDoc.AddSection();

            // Iterates through all sections in the document.
            foreach (IWSection section in letterDoc.Sections)
            {
                // Sets common page settings.
                section.PageSetup.Margins.Right = 40;
                section.PageSetup.PageSize = new SizeF(600, 790);

                // Configures the first section.
                if (sectionNumber == 1)
                {
                    // Sets page margins.
                    section.PageSetup.Margins.Top = 0;
                    section.PageSetup.Margins.Left = letterMarginLeft;

                    // Inserts XHTML content into the section body.
                    section.Body.InsertXHTML(letterSection1);

                    sectionNumber++;
                }
                // Configures subsequent sections.
                else
                {
                    // Sets header distance from the top of the page.
                    section.PageSetup.HeaderDistance = 45;

                    // Starts the section on a new page.
                    section.BreakCode = SectionBreakCode.NewPage;

                    // Restarts page numbering for the new section.
                    section.PageSetup.PageStartingNumber = 2;
                    section.PageSetup.RestartPageNumbering = true;

                    // Sets page margins.
                    section.PageSetup.Margins.Top = 0;
                    section.PageSetup.Margins.Left = letterMarginLeft - 70;

                    // Creates a paragraph in the section header.
                    IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();

                    // Adds a right-aligned tab stop.
                    paragraph.ParagraphFormat.Tabs.AddTab(560.0f, TabJustification.Right, TabLeader.NoLeader);

                    // Adds header text before the page number field.
                    paragraph.AppendText("\t Page ");

                    // Inserts a PAGE field into the header.
                    WField field = paragraph.AppendField("CurrentPageNumber", FieldType.FieldPage) as WField;

                    IEntity entity = field;

                    // Iterates through the field entities and applies formatting.
                    while (entity != null && entity.NextSibling != null)
                    {
                        // Formats text ranges within the field.
                        if (entity is WTextRange textRange)
                        {
                            textRange.CharacterFormat.FontSize = 22;
                            textRange.CharacterFormat.FontName = "Times New Roman";
                        }
                        // Stops when the field end marker is reached.
                        else if (entity is WFieldMark fieldMark &&
                                 fieldMark.Type == FieldMarkType.FieldEnd)
                        {
                            break;
                        }

                        // Moves to the next sibling entity.
                        entity = entity.NextSibling;
                    }

                    // Adds a body paragraph to the section.
                    paragraph = section.AddParagraph();

                    // Sets the page size.
                    section.PageSetup.PageSize = new SizeF(600, 790);

                    // Inserts XHTML content into the section body.
                    section.Body.InsertXHTML(letterSection2);
                }
            }
            // Returns the generated Word document.
            return letterDoc;
        }
    }
}