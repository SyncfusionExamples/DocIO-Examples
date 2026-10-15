# Build an AI‑Powered Word Document Translator Using the .NET Word Library

This repository provides an example of how to translate the content of a Word document **Azure OpenAI** and the **[.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO)**. The sample reads content from a Word document, translates the text into the specified language, and saves the translated content as a new Word document.

## Process Behind Word Content Translation

This sample demonstrates how to automate Word document translation by combining the document-processing capabilities of the **.NET Word Library (DocIO)** with the language translation capabilities of **Azure OpenAI**.

The workflow consists of the following steps:

1. Load and read the content of a Word document.
2. Extract text from the document.
3. Send the extracted content to Azure OpenAI with a translation prompt.
4. Translate the content into the target language.
5. Write the translated text back into a new Word document.
6. Save the translated document as a new .docx file.

## Prerequisites

Before running the sample, ensure that you have:

- An Azure OpenAI resource.
- A deployed chat model in Azure OpenAI.
- A valid Azure OpenAI API key.
- The Syncfusion DocIO NuGet package installed.

## Steps to Use the Sample

1. Open the application where the Syncfusion DocIO package is installed.
2. Replace the following placeholders in the code:
   - `Replace your Azure OpenAI key` with your Azure OpenAI API key.
   - `https://your-resource-name.openai.azure.com/` with your Azure OpenAI endpoint.
   - `your-model-name` with your deployed Azure OpenAI model name.
3. Run the application.
4. Save the translated document as a new .docx file.
5. Specify the target language.
6. The application translates the document content and saves it as a new Word document.

## Input

- Source Word document (`.docx`)
- Target translation language

## Output

A new Word document containing the translated content.