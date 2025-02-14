import React from 'react';
import './App.css';

function App() {
   // Function for converting Word to PDF
   async function convertToPDF() {
    // API endpoint for converting Word to PDF
    const url = 'http://localhost:5211/api/pdf/convertwordtopdf';

    // Get the selected file from the file input
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    var formData = new FormData();

    // If a file is selected, append it to the FormData object
    if (file) {
      formData.append('InputFile', file);
    }else{
      alert('Please select a file to convert.');
      return;
    }

    // Append various form data properties
    formData.append('Password', document.getElementById('password').value);
    formData.append('EmbedFontsInPDF', document.getElementById('embedFonts').checked);
    formData.append('EditablePDF', document.getElementById('EditablePDF').checked);
    formData.append('AutoDetectComplexScript', document.getElementById('autoDetectComplexScript').checked);
    formData.append('TaggedPDF', document.getElementById('TaggedPDF').checked);
    formData.append('HeadingsAsPdfBookmarks', document.getElementById('HeadingsAsPdfBookmarks').checked);
    formData.append('IncludeComments', document.getElementById('IncludeComments').checked);
    formData.append('IncludeRevisionsMarks', document.getElementById('IncludeRevisionsMarks').checked);
    formData.append('pdfConformanceLevel', parseInt(document.getElementById('pdfConformanceLevel').value));

    try {
      // Send a POST request to the server with the FormData
      const response = await fetch(url, {
        method: 'POST',
        body: formData
      });

      // Get the response as a Blob (binary data)
      const blob = await response.blob();
      const fileName = file.name.substring(0, file.name.lastIndexOf(".")) + ".pdf";

      // Create a link element to trigger the download
      const link = document.createElement('a');
      link.href = window.URL.createObjectURL(blob);
      link.download = fileName;

      // Append the link to the document body, trigger the download, and remove the link
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (error) {
      // Catch any errors that occur during the process and log an error message
      console.error('Error uploading PDF:', error);
    }
  }

  return (
    <div className="App">
    <div>
    {/* Input element to select the file */}
      <div className='header'>
        <h1 style={{width: 'auto', padding: '7px 15px', margin: '10px 0px', height: '39px'}}>
        Convert Word to PDF – Demo
        </h1>
        <input style={{width:'100%', marginLeft: '13px', padding: '10px'}} id='fileInput' type="file" accept=".doc,.docx,.rtf,.dot,.dotm,.dotx,.docm" />
      </div>
      <div style={{margin: '48px 10px'}}>
      <div className="properties-div">
      {/* Text input for password */}
        <label htmlFor="password">Password (if encrypted) : </label>
        <input type="text" id="password" placeholder='Enter password'/>
      </div>
      <div className="properties-div">
        <input type="checkbox" id="embedFonts"/>
        <label htmlFor="embedFonts">Embed the complete font information in the converted PDF document.</label>
      </div>
      <div className="properties-div">
        <input type="checkbox" checked id="EditablePDF"/>
        <label htmlFor="EditablePDF">Preserve Word form fields as editable PDF form fields.</label>
      </div>
      <div className="properties-div">
        <input type="checkbox" id="autoDetectComplexScript"/>
        <label htmlFor="autoDetectComplexScript">Detect complex script text present in the Word document.</label>
      </div>
      <div className="properties-div">
        <input type="checkbox" id="TaggedPDF"/>
        <label htmlFor="TaggedPDF">Convert the PDF document as tagged PDF (PDF/UA).</label>
      </div>
      <div className="properties-div">
        <input type="checkbox" checked id="HeadingsAsPdfBookmarks"/>
        <label htmlFor="HeadingsAsPdfBookmarks">Preserve the headings of the Word document as PDF bookmarks.</label>
      </div>
      <div className="properties-div">
        <input type="checkbox" checked id="IncludeComments"/>
        <label htmlFor="IncludeComments">Include comments from the Word document in the PDF.</label>
      </div>
      <div className="properties-div">
        <input type="checkbox" checked id="IncludeRevisionsMarks"/>
        <label htmlFor="IncludeRevisionsMarks">Include revision of tracked changes Word document in the PDF.</label>
      </div>
      <div className="properties-div">
       {/* PDF Conformance Level */}
        <label htmlFor="pdfConformanceLevel">PDF conformance level : </label>
        <select id="pdfConformanceLevel">
        <option value="0">None</option>
          <option value="1">Pdf_A1B</option>
          <option value="3">Pdf_A2B</option>
          <option value="4">Pdf_A3B</option>
          <option value="5">Pdf_A1A</option>
          <option value="6">Pdf_A2A</option>
          <option value="7">Pdf_A2U</option>
          <option value="8">Pdf_A3A</option>
          <option value="9">Pdf_A3U</option>
          <option value="10">Pdf_A4</option>
          <option value="11">Pdf_A4E</option>
          <option value="12">Pdf_A4F</option>
        </select>
      </div>
      {/* Button to trigger the conversion process */}
      <button className="convert-btn" onClick={convertToPDF}>Convert Word to PDF</button>
      </div>
    </div>
  </div>
  );
}

export default App;