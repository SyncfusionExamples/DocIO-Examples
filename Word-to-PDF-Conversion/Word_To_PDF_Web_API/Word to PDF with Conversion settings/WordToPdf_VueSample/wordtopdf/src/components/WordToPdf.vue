<template>
  <div class="App">
    <div class="header">
      <h1 style="width: auto;padding: 7px;margin: 10px 0px;height: 39px;">
        Convert Word to PDF – Demo
      </h1>
      <input style="margin-left: 5px;padding: 10px;width: 100%;" type="file" ref="fileInput" accept=".docx, .doc"  />
    </div>
    <br /><br />
    <div>
      <div class="properties-div">
        <label for="password">Password (if encrypted) : </label>
        <input type="text" v-model="password" placeholder="Enter password"/>
      </div>
      <div class="properties-div">
        <input type="checkbox" id="embedFonts" v-model="embedFonts" />
        <label for="embedFonts">Embed the complete font information in the converted PDF document.</label>
      </div>
      <div class="properties-div">
        <input type="checkbox" checked id="EditablePDF" v-model="EditablePDF">
        <label for="EditablePDF">Preserve Word form fields as editable PDF form fields.</label>
      </div>
      <div class="properties-div">
        <input type="checkbox" id="autoDetectComplexScript" v-model="autoDetectComplexScript">
        <label for="autoDetectComplexScript">Detect complex script text present in the Word document.</label>
      </div>
      <div class="properties-div">
        <input type="checkbox" id="TaggedPDF" v-model="TaggedPDF">
        <label for="TaggedPDF">Convert the PDF document as tagged PDF (PDF/UA).</label>
      </div>
      <div class="properties-div">
        <input type="checkbox" checked id="HeadingsAsPdfBookmarks" v-model="HeadingsAsPdfBookmarks">
        <label for="HeadingsAsPdfBookmarks">Preserve the headings of the Word document as PDF bookmarks.</label>
      </div>
      <div class="properties-div">
        <input type="checkbox" checked id="IncludeComments" v-model="IncludeComments">
        <label for="IncludeComments">Include comments from the Word document in the PDF.</label>
      </div>
      <div class="properties-div">
        <input type="checkbox" checked id="IncludeRevisionsMarks" v-model="IncludeRevisionsMarks">
        <label for="IncludeRevisionsMarks">Include revision of tracked changes Word document in the PDF.</label>
      </div>
      <div class="properties-div">
        <label for="pdfConformanceLevel">PDF conformance level : </label>
        <select v-model="pdfConformanceLevel">
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

      <button class="convert-btn" @click="convertToPDF">Convert Word to PDF</button>
    </div>
  </div>
</template>

<script>
// Exporting a default object with data and methods for Vue component
export default {
  data() {
    return {
      // Data properties for various form fields and settings
      password: '',                           
      embedFonts: false,                      
      EditablePDF: true,                      
      autoDetectComplexScript: false,         
      TaggedPDF: false,                       
      HeadingsAsPdfBookmarks: true,           
      IncludeComments: true,                 
      IncludeRevisionsMarks: true,           
      pdfConformanceLevel: "0",               
    };
  },
  methods: {
    async convertToPDF() {
      // API endpoint for converting Word to PDF
      const url = 'http://localhost:5211/api/pdf/convertwordtopdf';
      
      // Get the selected file from the file input using a ref
      const file = this.$refs.fileInput.files[0];
      const formData = new FormData();

      // If a file is selected, append it to the FormData object
      if (file) {
        formData.append('InputFile', file);
      }else{
        alert('Please select a file to convert.');
        return;
      }

      // Append various form data properties
      formData.append('Password', this.password);
      formData.append('EmbedFontsInPDF', this.embedFonts);
      formData.append('EditablePDF', this.EditablePDF);
      formData.append('AutoDetectComplexScript', this.autoDetectComplexScript);
      formData.append('TaggedPDF', this.TaggedPDF);
      formData.append('HeadingsAsPdfBookmarks', this.HeadingsAsPdfBookmarks);
      formData.append('IncludeComments', this.IncludeComments);
      formData.append('IncludeRevisionsMarks', this.IncludeRevisionsMarks);
      formData.append('pdfConformanceLevel', parseInt(this.pdfConformanceLevel));

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
  }
};
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
      .properties-div {
        margin: 10px;
      }
      .convert-btn{
        margin: 10px;
        padding: 4px;
      }
</style>
