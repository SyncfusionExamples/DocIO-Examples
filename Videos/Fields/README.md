# How to Work with Fields in Word Documents Using the .NET Word Library

This repository provides an example of how to work with fields in a Word document using the **Syncfusion .NET Word Library (DocIO)**. It demonstrates how to add field, update field, and unlink field in a Word document.

## Process behind Field Integration

This sample demonstrates how Word fields can be used to automate dynamic content in documents. Fields act as placeholders that display information such as page numbers, total pages, or dates, and they update automatically when the document changes.

Using the Syncfusion DocIO library, you can:

- Add fields like page numbers into headers or footers for consistent formatting.
- Update fields to refresh their values after edits or content changes.
- Unlink fields to convert them into static text when you no longer need dynamic updates.

## Steps to use the sample

1. Open the ASP.NET Core application where the Syncfusion DocIO package is installed.
2. Run the application and click the following buttons:
   - **AddField**: Creates a Word document with page number fields.
   - **UpdateField**: Updates all fields in the document.
   - **UnlinkField**: Converts date fields to static text.
