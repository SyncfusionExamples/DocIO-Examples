import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
// Define the AppComponent class
export class AppComponent {
  // Define a title property (optional)
  title = 'wordtopdf';

  // Inject HttpClient for making HTTP requests
  constructor(private http: HttpClient) {}

  // Initialize pdfConformanceLevel property
  pdfConformanceLevel: number = 0;

  // Method for converting Word to PDF
  async convertToPDF() {
    // Define the URL for the API endpoint
    const url = 'http://localhost:5211/api/pdf/convertwordtopdf';

    // Get the file input element and selected file
    const fileInput = (<HTMLInputElement>document.getElementById('fileInput'));
    const file = fileInput?.files?.[0];

    // Create a FormData object to hold the form data
    const formData = new FormData();

    // If a file is selected, append it to the FormData object
    if (file) {
      formData.append('InputFile', file);
    }else{
      alert('Please select a file to convert.');
      return;
    }

    // Append various form data properties
    formData.append('Password', (<HTMLInputElement>document.getElementById('password')).value);
    formData.append('EmbedFontsInPDF', (<HTMLInputElement>document.getElementById('embedFonts')).checked.toString());
    formData.append('EditablePDF', (<HTMLInputElement>document.getElementById('EditablePDF')).checked.toString());
    formData.append('AutoDetectComplexScript', (<HTMLInputElement>document.getElementById('autoDetectComplexScript')).checked.toString());
    formData.append('TaggedPDF', (<HTMLInputElement>document.getElementById('TaggedPDF')).checked.toString());
    formData.append('HeadingsAsPdfBookmarks', (<HTMLInputElement>document.getElementById('HeadingsAsPdfBookmarks')).checked.toString());
    formData.append('IncludeComments', (<HTMLInputElement>document.getElementById('IncludeComments')).checked.toString());
    formData.append('IncludeRevisionsMarks', (<HTMLInputElement>document.getElementById('IncludeRevisionsMarks')).checked.toString());
    formData.append('pdfConformanceLevel', this.pdfConformanceLevel.toString());

    try {
      // Send a POST request to the server with the FormData
      const response = await this.http.post(url, formData, { responseType: 'blob' }).toPromise();

      // Check if the response is a Blob
      if (response instanceof Blob) {
        // Process the Blob
        const blob = response;
        const fileName = file?.name.substring(0, file.name.lastIndexOf(".")) + ".pdf";

        // Create a link element to trigger the download
        const link = document.createElement('a');
        link.href = window.URL.createObjectURL(blob);
        link.download = fileName;

        // Append the link to the document body, trigger the download, and remove the link
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      } else {
        // Handle invalid response from server
        console.error('Error: Invalid response from server');
      }
    } catch (error) {
      // Handle errors during the process
      console.error('Error uploading PDF:', error);
    }
  }
}
