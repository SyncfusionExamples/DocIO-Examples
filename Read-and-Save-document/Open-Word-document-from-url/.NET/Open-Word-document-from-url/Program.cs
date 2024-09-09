using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Net;

namespace Open_Word_document_from_url
{
    class Program
    {
        static void Main(string[] args)
        {
            //Gets the document as stream.
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://www.syncfusion.com/downloads/support/directtrac/general/doc/Template235393797.docx");
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream stream = response.GetResponseStream();
            //Converts it to byte array
            byte[] buffer = ReadFully(stream, 32768);
            //Stores bytes into the memory stream.
            MemoryStream ms = new MemoryStream();
            ms.Write(buffer, 0, buffer.Length);
            ms.Seek(0, SeekOrigin.Begin);
            stream.Close();
            //Creates a new document.
            using (WordDocument document = new WordDocument())
            {
                //Opens the template document from the MemoryStream.
                document.Open(ms, FormatType.Doc);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        #region HelperMethods
        /// <summary>
        /// Reads the byte array from stream.
        /// </summary>
        public static byte[] ReadFully(Stream stream, int initialLength)
        {
            //When an unhelpful initial length has been passed, just use 32K.
            if (initialLength < 1)
            {
                initialLength = 32768;
            }
            byte[] buffer = new byte[initialLength];
            int read = 0;
            int chunk;
            while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
            {
                read += chunk;
                //After reaching the end of the buffer, check and see whether you can find any information.
                if (read == buffer.Length)
                {
                    int nextByte = stream.ReadByte();
                    //End of stream? Then, you are done.
                    if (nextByte == -1)
                    {
                        return buffer;
                    }
                    //Resize the buffer, put in the byte you have just read, and continue.
                    byte[] newBuffer = new byte[buffer.Length * 2];
                    Array.Copy(buffer, newBuffer, buffer.Length);
                    newBuffer[read] = (byte)nextByte;
                    buffer = newBuffer;
                    read++;
                }
            }
            //Buffer is now too big. Shrink it.
            byte[] ret = new byte[read];
            Array.Copy(buffer, ret, read);
            return ret;
        }
        #endregion
    }
}
