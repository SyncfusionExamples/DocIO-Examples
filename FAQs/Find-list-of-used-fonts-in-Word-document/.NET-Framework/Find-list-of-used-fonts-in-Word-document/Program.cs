using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace Find_list_of_used_fonts_in_Word_document
{
    internal class Program
    {
        static List<string> fontname = new List<string>();
        static void Main(string[] args)
        {
            //Open the template document.
            WordDocument wordDocument = new WordDocument(@"../../Data/Adventure.docx", FormatType.Docx);
            //Get the list of fonts used in current Word document.
            List<Font> usedFontsCollection = GetUsedFontsList(wordDocument);
            foreach (Font usedFont in usedFontsCollection)
            {
                Console.WriteLine(usedFont.Name);
            }
            //Save word document.
            wordDocument.Save("Output.docx");
            wordDocument.Close();
            Console.ReadKey();
            System.Diagnostics.Process.Start("Output.docx");
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
                //Gets the font from CharacterFormat or BreakCharacterFormat.
                font = HasProperty(item, "CharacterFormat") ? item.CharacterFormat.Font :
                    (HasProperty(item, "BreakCharacterFormat") ? item.BreakCharacterFormat.Font : null);

                //Gets the font from ListFormat.
                Font fontlist = GetUsedFontFromListFormat(item, usedFonts);
                if (fontlist != null)
                {
                    usedFonts.Add(fontlist);
                }

                //Adds the font name in the usedFont collection.
                if (font is Font && !(usedFonts.Find((match) => match.Name == font.Name && match.Italic == font.Italic && match.Bold == font.Bold) is Font))
                {
                    usedFonts.Add(font);
                }
            }
            return usedFonts;
        }
        /// <summary>
        /// Gets the fonts used in the List Format.
        /// </summary>
        /// <param name="item">current item</param>
        /// <param name="usedFonts">list of fonts</param>
        /// <returns></returns>
        private static Font GetUsedFontFromListFormat(dynamic item, List<Font> usedFonts)
        {
            if (item is WParagraph)
            {
                //if item is a paragraph can also include lists with own fonts
                if (item.ListFormat != null && item.ListFormat.CurrentListLevel != null)
                {
                    Font fontlist = item.ListFormat.CurrentListLevel.CharacterFormat.Font;
                    if (fontlist is Font && !(usedFonts.Find((match) => match.Name == fontlist.Name && match.Italic == fontlist.Italic && match.Bold == fontlist.Bold) is Font))
                    {
                        return fontlist;
                    }
                }
            }
            return null;
        }
        private static bool HasProperty(dynamic obj, string name)
        {
            Type objType = obj.GetType();
            if (objType == typeof(ExpandoObject))
            {
                return ((IDictionary<string, object>)obj).ContainsKey(name);
            }
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
