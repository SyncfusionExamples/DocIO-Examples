Imports System
Imports System.IO
Imports Syncfusion.DocIO
Imports Syncfusion.DocIO.DLS
Module Program
    Sub Main(args As String())
        'Loads a template document
        Dim document As New WordDocument(Path.GetFullPath("../../../Data/Destination.docx"), FormatType.Docx)
        'Gets the document to replace the text
        Dim replaceDocument As IWordDocument = New WordDocument(Path.GetFullPath("../../../Data/Template.docx"))
        'Replaces a particular text with another document
        document.Replace("<<PlaceHolder Text>>", replaceDocument, False, True, True)
        'Saves and closes the document
        document.Save(Path.GetFullPath("../../../Output/Sample.docx"), FormatType.Docx)
        document.Close()
    End Sub
End Module
