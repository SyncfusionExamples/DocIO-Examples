# Word to PDF Conversion API

## Overview

This repository contains a Web API application in ASP.NET to convert Word to PDF with conversion settings. You can refer to this to maintain the Web API on your side. Additionally, you can refer to the Blazor, Vue, JavaScript, React, and Angular repositories to learn how to invoke this Web API on different platforms.

## How to run the sample

1. Clone or Download the Repository

2. Build the Application

3. Run the API

This will start the API, and Swagger UI will be available at:

```
http://localhost:<port>/swagger
```

## Using the API

1. Open Swagger UI

Navigate to:

```
http://localhost:<port>/swagger
```

in your browser.

2. Upload a Word Document

- Use the `POST /convert` endpoint.
- Click **Try it out**.
- Select a Word document (`.docx`) in the **InputFile** field.
- Adjust the optional conversion settings as needed.
- Click **Execute** to convert the document.

3. Response

- The API will return a downloadable PDF file.
- If an error occurs, it will return an appropriate error message.

## Word to PDF Conversion Settings

The API supports the following optional settings:

- **Password** – Password to open a protected document.
- **EmbedFontsInPDF** – Embed fonts in the PDF (default: `false`).
- **EditablePDF** – Preserve form fields as editable fields (default: `true`).
- **AutoDetectComplexScript** – Detect complex script text (default: `false`).
- **TaggedPDF** – Convert to tagged PDF (default: `false`).
- **PdfConformanceLevel** – Specify the PDF conformance level.
- **HeadingsAsPdfBookmarks** – Preserve Word headings as bookmarks (default: `false`).
- **IncludeComments** – Include comments in the PDF (default: `false`).
- **IncludeRevisionsMarks** – Include revision marks in the PDF (default: `false`).
