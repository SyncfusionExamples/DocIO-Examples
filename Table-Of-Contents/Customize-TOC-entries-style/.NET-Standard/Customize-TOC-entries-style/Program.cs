using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Customize_TOC_entries_style
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                document.EnsureMinimal();
                document.LastSection.PageSetup.Margins.All = 72;
                WParagraph para = document.LastParagraph;
                para.AppendText("Essential DocIO - Table of Contents");
                para.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                para.ApplyStyle(BuiltinStyle.Heading4);
                para = document.LastSection.AddParagraph() as WParagraph;
                para = document.LastSection.AddParagraph() as WParagraph;
                //Insert TOC field in the Word document.
                TableOfContent toc = para.AppendTOC(1, 3);
                //Set the heading levels 1 to 3, to include in TOC.
                toc.LowerHeadingLevel = 1;
                toc.UpperHeadingLevel = 3;
                //Add content to the Word document with built-in heading styles.
                WSection section = document.LastSection;
                WParagraph newPara = section.AddParagraph() as WParagraph;
                newPara.AppendBreak(BreakType.PageBreak);
                AddHeading(section, BuiltinStyle.Heading1, "Document with built-in heading styles", "This is the built-in heading 1 style. This sample demonstrates the TOC insertion in a word document. Note that DocIO can insert TOC field in a word document. It can refresh or update TOC field by using UpdateTableOfContents method. MS Word refreshes the TOC field after insertion. Please update the field or press F9 key to refresh the TOC.");
                AddHeading(section, BuiltinStyle.Heading2, "Section 1", "This is the built-in heading 2 style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 1", "This is the built-in heading 3 style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 2", "This is the built-in heading 3 style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");
                //Add a new section to the Word document.
                section = document.AddSection() as WSection;
                section.PageSetup.Margins.All = 72;
                section.BreakCode = SectionBreakCode.NewPage;
                AddHeading(section, BuiltinStyle.Heading2, "Section 2", "This is the built-in heading 2 style. A document can contain any number of sections. Sections are used to apply same formatting for a group of paragraphs. You can insert sections by inserting section breaks.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 1", "This is the built-in heading 3 style. Each section contains any number of paragraphs. A paragraph is a set of statements that gives a meaning for the text.");
                AddHeading(section, BuiltinStyle.Heading3, "Paragraph 2", "This is the built-in heading 3 style. This demonstrates the paragraphs at the same level and style as that of the previous one. A paragraph can have any number formatting. This can be attained by formatting each text range in the paragraph.");

                //Access the TOC 1 style and update its formatting.
                IWParagraphStyle toc1style = document.AddParagraphStyle("TOC 1");
                toc1style.CharacterFormat.FontName = "Calibri";
                toc1style.CharacterFormat.FontSize = 14;
                toc1style.CharacterFormat.Bold = true;
                toc1style.CharacterFormat.Italic = true;
                toc1style.ParagraphFormat.AfterSpacing = 8;

                //Access the TOC 2 style and update its formatting.
                IWParagraphStyle toc2style = document.AddParagraphStyle("TOC 2");
                toc2style.CharacterFormat.FontName = "Calibri";
                toc2style.CharacterFormat.FontSize = 12;
                toc2style.ParagraphFormat.AfterSpacing = 5;
                toc2style.CharacterFormat.Bold = true;
                toc2style.ParagraphFormat.LeftIndent = 11;

                //Access the TOC 3 style and update its formatting.
                IWParagraphStyle toc3style = document.AddParagraphStyle("TOC 3"); ;
                toc3style.CharacterFormat.FontName = "Calibri";
                toc3style.CharacterFormat.FontSize = 12;
                toc3style.ParagraphFormat.AfterSpacing = 3;
                toc3style.CharacterFormat.Italic = true;
                toc3style.ParagraphFormat.LeftIndent = 22;

                //Update the table of contents.
                document.UpdateTableOfContents();
                //Save the file in the given path.
                Stream docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
        /// <summary>
        /// Add Heading and Apply BuiltinStyle. 
        /// </summary>
        private static void AddHeading(WSection section, BuiltinStyle builtinStyle, string headingText, string paragraghText)
        {
            WParagraph newPara = section.AddParagraph() as WParagraph;
            WTextRange text = newPara.AppendText(headingText) as WTextRange;
            newPara.ApplyStyle(builtinStyle);
            newPara = section.AddParagraph() as WParagraph;
            newPara.AppendText(paragraghText);
            section.AddParagraph();
        }
    }
    
}
