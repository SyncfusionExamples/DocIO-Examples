# Build an AI-Powered Word Document Summarizer Using the .NET Word Library

This repository provides an example of how to summarize the content of a Word document using **Azure OpenAI** and the **[.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) (DocIO)**. The sample reads text from a Word document, generates an AI-powered summary, and saves the summarized content as a new Word document.

## Process Behind Word Document Summarization

This sample demonstrates how to automate document summarization by combining the document-processing capabilities of the **.NET Word Library (DocIO)** with the natural language processing capabilities of **Azure OpenAI**.

The workflow consists of the following steps:

1. Load and read the content of a Word document.
2. Extract the document text using DocIO.
3. Send the extracted content to Azure OpenAI with a summarization prompt.
4. Generate a concise summary based on the specified number of sentences.
5. Create a new Word document containing the generated summary.
6. Save the summarized content as a new Word document.

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
4. Enter the full path of the Word document to summarize.
5. Specify the number of sentences required in the summary.
6. The application generates the summary and saves it as a new Word document.

## Input

- Source Word document (`.docx`)
- Desired summary length (number of sentences)

## Output

A new Word document containing the summarized content: