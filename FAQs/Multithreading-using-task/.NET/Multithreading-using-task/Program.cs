using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Multithreading_using_task
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
                tasks[i] = Task.Run(() => OpenAndSaveWordDocument());
            }
            //Ensure all tasks complete by waiting on each task.
            await Task.WhenAll(tasks);
        }

        //Open and save a Word document using multi-threading.
        static void OpenAndSaveWordDocument()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    // Save the Word document in the desired format
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + Guid.NewGuid().ToString() + ".docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
