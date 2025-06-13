import clr #import clr from pythonnet
import os

#Get the DLL and input file path in string
dll_path = r'D:\Create-document\Create-document\bin\Release\netstandard2.0\Create-document.dll' #Dll of application
docio_dll = r'C:\Users\Admin\.nuget\packages\syncfusion.docio.net.core\version\lib\netstandard2.0\Syncfusion.DocIO.Portable.dll'
officeChart_dll = r'C:\Users\Admin\.nuget\packages\syncfusion.officechart.net.core\version\lib\netstandard2.0\Syncfusion.OfficeChart.Portable.dll'
compression_dll =r'C:\Users\Admin\.nuget\packages\syncfusion.compression.net.core\version\lib\netstandard2.0\Syncfusion.Compression.Portable.dll'
license_dll = r'C:\Users\Admin\.nuget\packages\syncfusion.licensing\version\lib\netstandard2.0\Syncfusion.Licensing.dll'

# Verify and add paths
for path in [dll_path, docio_dll, officeChart_dll, compression_dll, license_dll]:
   if path not in os.sys.path:
       os.sys.path.append(path)
       
#load our dll file
try: 
   clr.AddReference(dll_path)
   clr.AddReference(docio_dll)
   clr.AddReference(officeChart_dll)
   clr.AddReference(compression_dll)
   clr.AddReference(license_dll)
   print ("Load success")
except Exception as e:
   print("Fail to load")

#import our DocIO class from Our C# namespace DocIOLibrary
from Create_document import CreateWordDocument

document = CreateWordDocument() #create our DocIO object

# Define file path
output_doc = r"Create-Word-document.docx" #path of result document

# Call CompareDocuments method
result = document.WordDocument(output_doc)

print("Document generated:", result) 
