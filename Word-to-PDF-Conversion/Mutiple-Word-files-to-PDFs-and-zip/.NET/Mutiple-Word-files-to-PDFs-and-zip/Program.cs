using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

//Creating new zip archive.
Syncfusion.Compression.Zip.ZipArchive zipArchive = new Syncfusion.Compression.Zip.ZipArchive();
//You can use CompressionLevel to reduce the size of the file.
zipArchive.DefaultCompressionLevel = Syncfusion.Compression.CompressionLevel.Best;

//Get the input Word documents from the folder.
string folderName = @"Data";
string[] inputFiles = Directory.GetFiles(folderName);
DirectoryInfo directoryInfo = new DirectoryInfo(folderName);
List<string> files = new List<string>();
FileInfo[] fileInfo = directoryInfo.GetFiles();
foreach (FileInfo fi in fileInfo)
{
    string name = Path.GetFileNameWithoutExtension(fi.Name);
    files.Add(name);
}

//Convert each Word documents to PDF.
for (int i = 0; i < inputFiles.Length; i++)
{
    //Output PDF file name.
    string outputFileName = files[i] + ".pdf";
    //Load an existing Word document.
    using (FileStream inputStream = new FileStream(inputFiles[i], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    {
        //Convert Word document to PDF.
        MemoryStream outputStream = ConvertWordToPDF(inputStream);
        //Add the converted PDF file into zip archive.
        zipArchive.AddItem(outputFileName, outputStream, true, Syncfusion.Compression.FileAttributes.Normal);
    }
}

//Zip file name and location.
FileStream zipStream = new FileStream(@"Output/Output.zip", FileMode.OpenOrCreate, FileAccess.ReadWrite);
zipArchive.Save(zipStream, true);
zipArchive.Close();


/// <summary>
/// Convert Word document to PDF.
/// </summary>
/// <param name="inputStream">Input Word document stream</param>
static MemoryStream ConvertWordToPDF(FileStream inputStream)
{
    MemoryStream outputStream = new MemoryStream();
    //Open an existing Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
    {
        //Initialize DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Convert Word document into PDF document
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
            {
                //Save the PDF.
                pdfDocument.Save(outputStream);
                outputStream.Position = 0;
            }
        }
    }
    return outputStream;
}