using Syncfusion.Office.Markdown;
using System.Text;

namespace Create_markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Markdown document
            MarkdownDocument doc = new MarkdownDocument();

            // ── HEADING 1 – Document title
            MdParagraph h1 = doc.AddParagraph();
            h1.ApplyParagraphStyle("Heading 1");
            h1.AddTextRange().Text = "The Giant Panda";

            // ── Image section
            MdParagraph imgPara = doc.AddParagraph();
            MdPicture pandaImage = new MdPicture();
            imgPara.Inlines.Add(pandaImage);
            // Set image source and alternative text
            pandaImage.Url = "https://cdn.syncfusion.com/content/images/company-logos/Syncfusion_Logo_Image.png";
            pandaImage.AltText = "Giant Panda eating bamboo";
            // Add image caption
            MdParagraph imgCaption = doc.AddParagraph();
            MdTextRange captionText = imgCaption.AddTextRange();
            captionText.Text = "Figure 1: Syncfusion logo used for demonstration purposes.";
            captionText.TextFormat.Italic = true;

            // ── Introduction paragraph 
            MdParagraph intro = doc.AddParagraph();
            intro.AddTextRange().Text = "The giant panda, which only lives in ";
            // Add bold formatted text
            MdTextRange boldChina = intro.AddTextRange();
            boldChina.Text = "China";
            boldChina.TextFormat.Bold = true;
            intro.AddTextRange().Text = " outside of captivity, has captured the hearts of people of all ages across the globe. ";

            // ── Blockquote – Quick Fact 
            MdParagraph quickFact1 = doc.AddParagraph();
            quickFact1.HasBlockquote = true;
            quickFact1.BlockQuoteLevel = 1;
            quickFact1.AddTextRange().Text = "🐼 Quick Fact: The estimated number of giant pandas in the wild varies between 1,500 and 3,000.";
            // Add horizontal line
            doc.AddThematicBreak();

            // ── HEADING 2 – Intriguing Mysteries
            MdParagraph h2Mysteries = doc.AddParagraph();
            h2Mysteries.ApplyParagraphStyle("Heading 2");
            h2Mysteries.AddTextRange().Text = "Intriguing Giant Panda Mysteries";

            // ── PARAGRAPH WITH MIXED FORMATTING
            MdParagraph mysteriesPara = doc.AddParagraph();
            mysteriesPara.AddTextRange().Text = "While most adore their fluffy fur and round heads, others are fascinated by the many mysteries of the giant panda. " +
                "Did you know that the giant panda may actually be a raccoon, they have an ";
            MdTextRange boldItalicThumb = mysteriesPara.AddTextRange();
            boldItalicThumb.Text = "opposable pseudo thumb";
            boldItalicThumb.TextFormat.Bold = true;
            boldItalicThumb.TextFormat.Italic = true;
            mysteriesPara.AddTextRange().Text = ", and that they're technically a ";
            // Strikethrough text
            MdTextRange strikeOld = mysteriesPara.AddTextRange();
            strikeOld.Text = "tennis ball";
            strikeOld.TextFormat.StrikeThrough = true;
            mysteriesPara.AddTextRange().Text = "carnivore even though their diet is primarily vegetarian?";

            // ── TABLE – Bear vs Red Panda characteristics
            MdParagraph tableIntro = doc.AddParagraph();
            tableIntro.AddTextRange().Text = "The table below lists the main characteristics the giant panda shares with bears and red pandas:";
            // Create table with column alignment
            MdTable compTable = doc.AddTable();
            compTable.ColumnAlignments.Add(MdColumnAlignment.Left);
            compTable.ColumnAlignments.Add(MdColumnAlignment.Center);
            compTable.ColumnAlignments.Add(MdColumnAlignment.Center);

            // Header row
            MdTableRow compHeader = compTable.AddTableRow();
            MdTextRange hCharacteristic = new MdTextRange { Text = "Characteristic" };
            hCharacteristic.TextFormat.Bold = true;
            compHeader.AddTableCell().Items.Add(hCharacteristic);
            MdTextRange hBear = new MdTextRange { Text = "Bear" };
            hBear.TextFormat.Bold = true;
            compHeader.AddTableCell().Items.Add(hBear);
            MdTextRange hRedPanda = new MdTextRange { Text = "Red Panda" };
            hRedPanda.TextFormat.Bold = true;
            compHeader.AddTableCell().Items.Add(hRedPanda);

            // Add data rows
            string[][] compRows = {
            new[] { "Shape",       "✅", "❌" },
            new[] { "Diet",        "❌", "✅" },
            new[] { "Size",        "✅", "❌" },
            new[] { "Paws",        "✅", "✅" },
            new[] { "Cat-like eyes","❌","✅" },
            };
            foreach (string[] rowData in compRows)
            {
                MdTableRow row = compTable.AddTableRow();
                row.AddTableCell().Items.Add(new MdTextRange { Text = rowData[0] });
                row.AddTableCell().Items.Add(new MdTextRange { Text = rowData[1] });
                row.AddTableCell().Items.Add(new MdTextRange { Text = rowData[2] });
            }

            // ── Nested blockquote
            MdParagraph quickFact2 = doc.AddParagraph();
            quickFact2.HasBlockquote = true;
            quickFact2.BlockQuoteLevel = 1;
            quickFact2.AddTextRange().Text = "🌿 Quick Fact: As the seasons change, the giant panda prefers different species and parts of bamboo.";

            MdParagraph nestedQuote = doc.AddParagraph();
            nestedQuote.HasBlockquote = true;
            nestedQuote.BlockQuoteLevel = 2;
            nestedQuote.AddTextRange().Text = "For comparison, humans eat about 2 kilograms (5 pounds) of food a day.";
            doc.AddThematicBreak();

            // ── HEADING 2 – Other Fun Facts (Unordered / Bulleted List)
            MdParagraph h2Facts = doc.AddParagraph();
            h2Facts.ApplyParagraphStyle("Heading 3");
            h2Facts.AddTextRange().Text = "Other Fun Giant Panda Facts";

            // ── Unordered (Bulleted) List — ListValue must be "- " for each item
            MdParagraph fact1 = doc.AddParagraph();
            fact1.ListFormat = new MdListFormat();
            fact1.ListFormat.IsNumbered = false;
            fact1.ListFormat.ListLevel = 0;
            fact1.ListFormat.ListValue = "- ";
            fact1.AddTextRange().Text = "Unlike other bears, giant pandas do not hibernate during winter.";

            MdParagraph fact2 = doc.AddParagraph();
            fact2.ListFormat = new MdListFormat();
            fact2.ListFormat.IsNumbered = false;
            fact2.ListFormat.ListLevel = 0;
            fact2.ListFormat.ListValue = "- ";
            fact2.AddTextRange().Text = "A newborn giant panda is blind and looks like a tiny, pink, hairless mouse.";
            doc.AddThematicBreak();

            // ── HEADING 2 – Conservation Steps (Ordered / Numbered List)
            MdParagraph h2Conservation = doc.AddParagraph();
            h2Conservation.ApplyParagraphStyle("Heading 2");
            h2Conservation.AddTextRange().Text = "Conservation Steps";

            MdParagraph conservationIntro = doc.AddParagraph();
            conservationIntro.AddTextRange().Text = "Key steps taken to protect giant pandas in the wild:";

            // ── Ordered (Numbered) List — ListValue must carry the sequential number e.g. "1. ", "2. ", "3. "
            MdParagraph cStep1 = doc.AddParagraph();
            cStep1.ListFormat = new MdListFormat();
            cStep1.ListFormat.IsNumbered = true;
            cStep1.ListFormat.ListLevel = 0;
            cStep1.ListFormat.NumberedListMarker = "1.";
            cStep1.ListFormat.ListValue = "1. ";
            cStep1.AddTextRange().Text = "Protect and expand bamboo forest reserves across Sichuan and Shaanxi.";

            MdParagraph cStep2 = doc.AddParagraph();
            cStep2.ListFormat = new MdListFormat();
            cStep2.ListFormat.IsNumbered = true;
            cStep2.ListFormat.ListLevel = 0;
            cStep2.ListFormat.NumberedListMarker = "1.";
            cStep2.ListFormat.ListValue = "2. ";
            cStep2.AddTextRange().Text = "Establish wildlife corridors to connect isolated panda habitats.";
            doc.AddThematicBreak();

            // ── HEADING 2 – Research Checklist (Task List)
            MdParagraph h2Tasks = doc.AddParagraph();
            h2Tasks.ApplyParagraphStyle("Heading 2");
            h2Tasks.AddTextRange().Text = "Researcher's Checklist";
            // Task 1 — checked (completed)
            MdParagraph task1 = doc.AddParagraph();
            task1.TaskItemProperties = new MdTaskProperties { IsChecked = true };
            task1.AddTextRange().Text = "Review existing population data";
            // Task 2 — unchecked (pending)
            MdParagraph task2 = doc.AddParagraph();
            task2.TaskItemProperties = new MdTaskProperties { IsChecked = false };
            task2.AddTextRange().Text = "Deploy GPS tracking collars on selected individuals";
            doc.AddThematicBreak();

            // ── HEADING 2 – Text Formatting Showcase
            MdParagraph h2Format = doc.AddParagraph();
            h2Format.ApplyParagraphStyle("Heading 2");
            h2Format.AddTextRange().Text = "Scientific Classification";
            MdParagraph classificationPara = doc.AddParagraph();
            classificationPara.AddTextRange().Text = "Species: ";
            MdTextRange speciesName = classificationPara.AddTextRange();
            speciesName.Text = "Ailuropoda melanoleuca";
            speciesName.TextFormat.Italic = true;
            classificationPara.AddTextRange().Text = " | Family: ";
            MdTextRange family = classificationPara.AddTextRange();
            family.Text = "Ursidae";
            family.TextFormat.Bold = true;
            classificationPara.AddTextRange().Text = " | Status: ";
            MdTextRange status = classificationPara.AddTextRange();
            status.Text = "Vulnerable";
            status.TextFormat.Bold = true;
            status.TextFormat.Italic = true;
            classificationPara.AddTextRange().Text = " (IUCN Red List)";

            // ── Subscript / Superscript example
            MdParagraph formulaPara = doc.AddParagraph();
            formulaPara.AddTextRange().Text = "Cellulose formula: (C";
            MdTextRange sub6 = formulaPara.AddTextRange();
            sub6.Text = "6";
            sub6.TextFormat.SubSuperScriptType = MdSubSuperScript.SubScript;
            formulaPara.AddTextRange().Text = "H";
            MdTextRange sub10 = formulaPara.AddTextRange();
            sub10.Text = "10";
            sub10.TextFormat.SubSuperScriptType = MdSubSuperScript.SubScript;
            formulaPara.AddTextRange().Text = "O";
            MdTextRange sub5 = formulaPara.AddTextRange();
            sub5.Text = "5";
            sub5.TextFormat.SubSuperScriptType = MdSubSuperScript.SubScript;
            formulaPara.AddTextRange().Text = ")";
            MdTextRange sup = formulaPara.AddTextRange();
            sup.Text = "n";
            sup.TextFormat.SubSuperScriptType = MdSubSuperScript.SuperScript;
            formulaPara.AddTextRange().Text = " — the primary component of bamboo that pandas struggle to digest.";
            doc.AddThematicBreak();

            // ── HEADING 2 – Code Block (Taxonomy lookup example)
            MdParagraph h2Code = doc.AddParagraph();
            h2Code.ApplyParagraphStyle("Heading 2");
            h2Code.AddTextRange().Text = "Panda Population Tracker — Code Sample";

            MdParagraph codeIntro = doc.AddParagraph();
            codeIntro.AddTextRange().Text = "The following C# snippet demonstrates how a simple panda population record might be modelled:";

            MdCodeBlock codeBlock = doc.AddCodeBlock();
            codeBlock.IsFencedCode = true;
            codeBlock.Lines.Add("public class GiantPanda");
            codeBlock.Lines.Add("{");
            codeBlock.Lines.Add("    public string Name        { get; set; }");
            codeBlock.Lines.Add("    public double WeightKg    { get; set; }");
            codeBlock.Lines.Add("    public string Habitat     { get; set; }");
            codeBlock.Lines.Add("    public bool   IsEndangered { get; set; }");
            codeBlock.Lines.Add("");
            codeBlock.Lines.Add("    public override string ToString()");
            codeBlock.Lines.Add("        => $\"{Name} ({WeightKg} kg) — Habitat: {Habitat}\";");
            codeBlock.Lines.Add("}");

            MdParagraph codeDesc = doc.AddParagraph();
            codeDesc.AddTextRange().Text = "Inline usage: instantiate with ";
            MdTextRange inlineCode = codeDesc.AddTextRange();
            inlineCode.Text = "new GiantPanda { Name = \"Mei Xiang\", WeightKg = 124.7 }";
            inlineCode.TextFormat.CodeSpan = true;
            codeDesc.AddTextRange().Text = ".";
            doc.AddThematicBreak();

            // ── HEADING 2 – Hyperlinks & References
            MdParagraph h2Links = doc.AddParagraph();
            h2Links.ApplyParagraphStyle("Heading 2");
            h2Links.AddTextRange().Text = "Further Reading";

            MdParagraph linksIntro = doc.AddParagraph();
            linksIntro.AddTextRange().Text = "Explore more about giant pandas through these authoritative sources:";

            // ── Bulleted list with hyperlinks — 3 items
            string[][] linkData = { new[] { "Panda Info", "https://example.com/panda" },
                                new[] { "Documentation", "https://example.com/docs" },
                                new[] { "API Reference", "https://example.com/api" }};
            foreach (string[] link in linkData)
            {
                MdParagraph linkItem = doc.AddParagraph();
                linkItem.ListFormat = new MdListFormat { IsNumbered = false, ListLevel = 0, ListValue = "- " };
                MdHyperlink hl = linkItem.AddHyperlink();
                hl.DisplayText = link[0];
                hl.Url = link[1];
                hl.ScreenTip = link[0];
            }
            doc.AddThematicBreak();

            // ── Footer
            MdParagraph footer = doc.AddParagraph();
            MdTextRange footerText = footer.AddTextRange();
            footerText.Text = "Document generated using Syncfusion Markdown library.";
            footerText.TextFormat.Italic = true;

            // ── SAVE MARKDOWN FILE LOCALLY
            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/CreateMarkdown.md"), FileMode.Create, FileAccess.Write))
            {
                // Convert document to markdown text
                string mdText = doc.GetMarkdownText();
                byte[] bytes = Encoding.UTF8.GetBytes(mdText);
                // Write content into file stream
                outputFileStream.Write(bytes, 0, bytes.Length);
            }
        }
    }
}
