Identify and manage file formats not supported by DocIO
----------------------------------------

DocIO supports Word 97-2003 and later versions, handling major Microsoft Word file formats, including DOC, DOCX, RTF, DOT, DOTX, and DOCM. If you use an older or unsupported format, convert it to a supported format to avoid errors. Otherwise, you will get this exception: **This file format is not supported**.

**To run this example**

*   Download this project to a location in your disk.
*   Create a folder **Data** parallelly to .sln file.
*   Add the input documents to the "Data" folder.
*   Run the application.

The console output will indicate whether the provided input document is supported or unsupported. If the document is unsupported, convert it to a format supported by DocIO.
