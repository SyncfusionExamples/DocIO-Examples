using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Text;

namespace Convert_RTF_String_To_HTML_String
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string rtfString = @"{\rtf1\ansi\deff0 {\fonttbl {\f0 Times New Roman;}}\f0\fs60 Hello World!}";
            byte[] bytes = Encoding.ASCII.GetBytes(rtfString);
            MemoryStream streamRTF = new MemoryStream(bytes);
            WordDocument document = new WordDocument(streamRTF, FormatType.Rtf);
            MemoryStream ms = new MemoryStream();
            document.Save(ms, FormatType.Html);
            document.Close();
            ms.Position = 0;
            streamRTF.Close();
            string htmlString = Encoding.ASCII.GetString(ms.ToArray());
			Console.WriteLine(htmlString);
            ms.Close();
        }
    }
}