const fs = require('fs');
const FormData = require('form-data');
const axios = require('axios');
const https = require('https');

const filePath1 = "Input.docx"; // Replace with the actual file path
const fileData1 = fs.readFileSync(filePath1);

const formData = new FormData();
formData.append('file', fileData1, {
  filename: 'Input.docx',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
});

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

axios.post('http://localhost:5083/api/docio/OpenAndResave', formData, {
  headers: formData.getHeaders(),
  responseType: 'arraybuffer', // Ensure response is treated as a binary file
  httpsAgent,
})
  .then(response => {
    console.log('File successfully processed');
    fs.writeFileSync('ResavedDocument.docx', response.data);
  })
  .catch(error => {
    console.error('Error:', error.message);
  });
