using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Multithreading_using_tasks
{
    class MultiThreading
    {
        //Indicates the number of threads to be create.
        private const int TaskCount = 1000;
        public static async Task Main()
        {
            //Create an array of tasks based on the TaskCount.
            Task[] tasks = new Task[TaskCount];
            for (int i = 0; i < TaskCount; i++)
            {
                tasks[i] = Task.Run(() => ConvertWordToPDF());
            }
            //Ensure all tasks complete by waiting on each task.
            await Task.WhenAll(tasks);
        }

        //Convert a Word document to PDF using multi-threading.
        static void ConvertWordToPDF()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load an existing Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Create an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert Word document to PDF.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
                        {
                            //Save the PDF document.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + Guid.NewGuid().ToString() + ".pdf"), FileMode.Create, FileAccess.Write))
                            {
                                pdfDocument.Save(outputFileStream);
                            }
                        }
                    }
                }
            }
        }
    }
}
