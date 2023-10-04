using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace Find_used_fonts_in_word_document
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the template document.
            WordDocument wordDocument = new WordDocument(@"../../Data/Adventure.docx", FormatType.Docx);
            //Get the list of fonts used in current Word document.
            List<Font> usedFontsCollection = GetUsedFontsList(wordDocument);
            Console.WriteLine("List of fonts used in the Word document:");
            foreach (Font usedFont in usedFontsCollection)
            {
                Console.WriteLine("Font name : " + usedFont.Name + "; Bold : " + usedFont.Bold + "; Italic : " + usedFont.Italic);
            }
            //Save word document.
            wordDocument.Save("Output.docx");
            wordDocument.Close();
            Console.WriteLine("Press enter key to exit.");
            Console.ReadKey();
        }
        /// <summary>
        /// Get the List of fonts which is used in current Word document.
        /// </summary>
        /// <param name="document">Word document</param>
        /// <returns>Used font collection</returns>
        private static List<Font> GetUsedFontsList(WordDocument document)
        {
            List<Font> usedFonts = new List<Font>();
            Font font = null;
            //Visits all document entities.
            foreach (dynamic item in document.Visit())
            {
                // Gets the font from CharacterFormat or BreakCharacterFormat.
                font = HasProperty(item, "CharacterFormat") ? item.CharacterFormat.Font :
                    (HasProperty(item, "BreakCharacterFormat") ? item.BreakCharacterFormat.Font : null);

                AddToFontsCollection(font, usedFonts);

                //Gets the font from ListFormat.
                GetUsedFontFromListFormat(item, usedFonts);
            }
            return usedFonts;
        }
        /// <summary>
        /// Add the font to the collection.
        /// </summary>
        /// <param name="font">Font to add.</param>
        /// <param name="usedFonts">Collection of fonts.</param>
        private static void AddToFontsCollection(Font font, List<Font> usedFonts)
        {
            // Check whether the font is different in name and font style properties and add it to the collection.
            if (font != null && font is Font &&
                !(usedFonts.Find((match) => match.Name == font.Name && match.Italic == font.Italic && match.Bold == font.Bold) is Font))
            {
                usedFonts.Add(font);
            }
        }
        /// <summary>
        /// Gets the fonts used in the List Format.
        /// </summary>
        /// <param name="item">Current item.</param>
        /// <param name="usedFonts">Collection of fonts.</param>
        /// <returns></returns>
        private static void GetUsedFontFromListFormat(dynamic item, List<Font> usedFonts)
        {
            if (item is WParagraph)
            {
                //if item is a paragraph then get the font from list format.
                if (item.ListFormat != null && item.ListFormat.CurrentListLevel != null)
                {
                    Font font = item.ListFormat.CurrentListLevel.CharacterFormat.Font;
                    AddToFontsCollection(font, usedFonts);
                }
            }
        }
        /// <summary>
        /// Check whether the property is available in the object.
        /// </summary>
        /// <param name="obj">Current object.</param>
        /// <param name="name">Property to check.</param>
        /// <returns>True, if object has property. Otherwise, false.</returns>
        private static bool HasProperty(dynamic obj, string name)
        {
            Type objType = obj.GetType();
            return objType.GetProperty(name) != null;
        }
    }

    #region ExtendedClass
    /// <summary>
    /// DocIO extension class.
    /// </summary>
    public static class DocIOExtensions
    {
        public static IEnumerable<IEntity> Visit(this ICompositeEntity entity)
        {
            var entities = new Stack<IEntity>(new IEntity[] { entity });
            while (entities.Count > 0)
            {
                var e = entities.Pop();
                yield return e;
                if (e is ICompositeEntity)
                {
                    foreach (IEntity childEntity in ((ICompositeEntity)e).ChildEntities)
                    {
                        entities.Push(childEntity);
                    }
                }
            }
        }
    }
    #endregion
}
