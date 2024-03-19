using System.Reflection;

namespace Use_alternate_font_without_installing
{
    public class HelperService
    {
        /// <summary>
        /// Gets the embedded font stream.
        /// </summary>
        /// <param name="fontName">Represent the name of the font stream.</param>
        /// <returns>Returns the font stream of given font name, if it is embedded. Otherwise returns null.</returns>
        public Stream GetFontStream(string fontName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            //fontName = fontName.ToLower() + ".ttf";
            foreach (string resourceName in assembly.GetManifestResourceNames())
            {
                if (resourceName.ToLower().EndsWith(fontName))
                {
                    Stream fontStream = assembly.GetManifestResourceStream(resourceName);
                    if (fontStream != null)
                    {
                        fontStream.Position = 0;
                        return fontStream;
                    }
                }
            }
            return null;
        }
    }
}
