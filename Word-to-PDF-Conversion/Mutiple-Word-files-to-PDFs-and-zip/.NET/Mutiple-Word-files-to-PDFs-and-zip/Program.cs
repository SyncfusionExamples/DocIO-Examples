using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO.Compression;
using Syncfusion.Compression.Zip;

//Creating new zip archive
Syncfusion.Compression.Zip.ZipArchive zipArchive = new Syncfusion.Compression.Zip.ZipArchive();
//You can use CompressionLevel to reduce the size of the file.
zipArchive.DefaultCompressionLevel = Syncfusion.Compression.CompressionLevel.Best;

//Get the input files from the folder
string folderName = @"../../../InputDocuments";
string[] inputFiles = Directory.GetFiles(folderName);
DirectoryInfo directoryInfo = new DirectoryInfo(folderName);
List<string> files = new List<string>();
FileInfo[] fileInfo = directoryInfo.GetFiles();
foreach (FileInfo fi in fileInfo)
{
    string name = Path.GetFileNameWithoutExtension(fi.Name);
    files.Add(name);
}

//Converts each Word documents to PDF documents
for (int i = 0; i < inputFiles.Length; i++)
{
    //PDF file name
    string outputFileName = files[i] + ".pdf";
    //Loads an existing Word document
    using (FileStream inputStream = new FileStream(inputFiles[i], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    {
        //Converts Word document to PDF document
        MemoryStream outputStream = ConvertWordToPDF(inputStream);
        //Add the converted PDF file to zip
        zipArchive.AddItem(outputFileName, outputStream, true, Syncfusion.Compression.FileAttributes.Normal);
    }
}

//Zip file name and location
FileStream zipStream = new FileStream(@"../../../OutputPDFs.zip", FileMode.OpenOrCreate, FileAccess.ReadWrite);
zipArchive.Save(zipStream, true);
zipArchive.Close();


/// <summary>
/// Convert Word document to PDF
/// </summary>
/// <param name="inputStream">Input Word document stream</param>
static MemoryStream ConvertWordToPDF(FileStream inputStream)
{
    using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
    {
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Converts Word document into PDF document
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
            {
                MemoryStream outputStream = new MemoryStream();
                pdfDocument.Save(outputStream);
                outputStream.Position = 0;
                //Closes the instance of PDF document object
                pdfDocument.Close();
                return outputStream;
            }
        }
    }
}