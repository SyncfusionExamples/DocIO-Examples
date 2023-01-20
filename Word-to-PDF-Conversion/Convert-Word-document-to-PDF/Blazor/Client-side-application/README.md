convert Word document to PDF in Blazor WebAssembly (WASM)
---------------------------------------------------------

This application shows how to convert Word document to PDF in Blazor WebAssembly (WASM).

Steps to convert Word document to PDF in Blazor WebAssembly (WASM)
------------------------------------------------------------------

1. Create a new C# Blazor WebAssembly App in Visual Studio.  
2. Install the [Syncfusion.DocIORenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIORenderer.Net.Core) NuGet package as a reference to your Blazor application from [NuGet.org](https://www.nuget.org/).  
3. Install the [SkiaSharp.NativeAssets.WebAssembly](https://www.nuget.org/packages/SkiaSharp.NativeAssets.WebAssembly) NuGet package as a reference to your Blazor application from [NuGet.org](https://www.nuget.org/).  
4. Add the following ItemGroup tag in the [Blazor WASM csproj](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/Convert-Word-to-PDF.csproj) file.

```xml
<ItemGroup>
    <NativeFileReference Include="$(SkiaSharpStaticLibraryPath)\2.0.23\*.a" />
</ItemGroup>
```

5. Create a [razor](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/Pages/DocIO.razor) file named as DocIO under the **Pages** folder and add the namespaces in the file.
6. Add the code to create a button.
7. Create a new async method with the name WordToPDF and include the code sample to convert a Word document to PDF.
8. Create a [class](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/FileUtils.cs) file with FileUtils name and add the code to invoke the JavaScript action to download the file in the browser.
9. Add the JavaScript function in the [Index.html](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/wwwroot/index.html) file present under the **wwwroot** folder.
10. Add the code sample in the [razor](https://github.com/SyncfusionExamples/DocIO-Examples/blob/main/Word-to-PDF-Conversion/Convert-Word-document-to-PDF/Blazor/Client-side-application/Convert-Word-to-PDF/Shared/NavMenu.razor) file of the Navigation menu in the **Shared** folder.
11. Rebuild the solution.
12. Run the application.
