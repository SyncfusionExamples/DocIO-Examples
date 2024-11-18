Convert Word document to PDF in Blazor WebAssembly (WASM)
---------------------------------------------------------

This example illustrates how to convert Word document to PDF in Blazor WebAssembly (WASM).

Steps to convert Word document to PDF in Blazor WebAssembly (WASM)
------------------------------------------------------------------

1. Create a new C# Blazor WebAssembly App in Visual Studio.  
2. Install the [Syncfusion.DocIORenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIORenderer.Net.Core) NuGet package as a reference to your Blazor application from [NuGet.org](https://www.nuget.org/).  
3. Install the [SkiaSharp.Views.Blazor v2.88.8](https://www.nuget.org/packages/SkiaSharp.Views.Blazor/2.88.8) NuGet package as a reference to your Blazor application from [NuGet.org](https://www.nuget.org/).  

> **Note:** Install this `wasm-tools` and `wasm-tools-net6` by using `dotnet workload install wasm-tools` and `dotnet workload install wasm-tools-net6` commands in your command prompt respectively, while facing issues related to skiasharp, during runtime.

4. Create a [razor](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/Pages/DocIO.razor) file named as DocIO under the **Pages** folder and add the namespaces in the file.
5. Add the code to create a button.
6. Create a new async method with the name WordToPDF and include the code sample to convert a Word document to PDF.
7. Create a [class](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/FileUtils.cs) file with FileUtils name and add the code to invoke the JavaScript action to download the file in the browser.
8. Add the JavaScript function in the [Index.html](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/wwwroot/index.html) file present under the **wwwroot** folder.
9. Add the code sample in the [razor](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/Shared/NavMenu.razor) file of the Navigation menu in the **Shared** folder.
10. Rebuild the solution.
11. Run the application.
