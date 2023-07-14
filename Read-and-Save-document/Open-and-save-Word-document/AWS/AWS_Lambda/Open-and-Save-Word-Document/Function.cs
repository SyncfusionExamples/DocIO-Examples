using Amazon.Lambda.Core;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Reflection;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Open_and_Save_Word_Document;

public class Function
{
    
    /// <summary>
    /// A simple function that takes a string and does a ToUpper
    /// </summary>
    /// <param name="input"></param>
    /// <param name="context"></param>
    /// <returns></returns>
    public string FunctionHandler(string input, ILambdaContext context)
    {
        string filePath = Path.GetFullPath(@"Data/Input.docx");
        //Load the file from the disk.
        using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
            {
                //Access the section in a Word document.
                IWSection section = document.Sections[0];
                //Add new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ParagraphFormat.FirstLineIndent = 36;
                paragraph.BreakCharacterFormat.FontSize = 12f;
                //Add new text to the paragraph.
                IWTextRange textRange = paragraph.AppendText("In 2000, AdventureWorks Cycles bought a small manufacturing plant, Importadores Neptuno, located in Mexico. Importadores Neptuno manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the Bothell location for final product assembly. In 2001, Importadores Neptuno, became the sole manufacturer and distributor of the touring bicycle product group.") as IWTextRange;
                textRange.CharacterFormat.FontSize = 12f;
                //Save the Word document.
                MemoryStream stream = new MemoryStream();
                document.Save(stream,FormatType.Docx);
                return Convert.ToBase64String(stream.ToArray());
            }
        }            
    }
}
