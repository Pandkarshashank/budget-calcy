const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs')

const app = express();

// Middleware to parse form data
app.use(bodyParser.urlencoded({ extended: true }));

// Serve the HTML form
app.get('/', (req, res) => {
  res.sendFile(__dirname + "/home.html");
});
app.get('/submit',(req,res)=>{
    res.sendFile(__dirname + "/index.html")
})
// Handle form submission
app.post('/submit', (req, res) => {
    const formData = req.body;
  
    // Load existing workbook or create a new one
    let workbook;
    if (fs.existsSync('data.xlsx')) {
      workbook = xlsx.readFile('data.xlsx');
    } else {
      workbook = xlsx.utils.book_new();
    }
  
    // Check if the workbook is empty
    const sheetNames = workbook.SheetNames;
    if (sheetNames.length === 0) {
      // Create a new worksheet and add the form data with the title row
      const newWorksheet = xlsx.utils.json_to_sheet([formData], { header: Object.keys(formData) });
      xlsx.utils.book_append_sheet(workbook, newWorksheet, 'Form Data');
    } else {
      // Check if the "Form Data" worksheet already exists
      const worksheet = workbook.Sheets['Form Data'];
      if (worksheet) {
        // Append the form data without adding the title row
        const jsonData = xlsx.utils.sheet_to_json(worksheet);
        jsonData.push(formData);
        const updatedWorksheet = xlsx.utils.json_to_sheet(jsonData);
        workbook.Sheets['Form Data'] = updatedWorksheet;
      } else {
        // Add the form data to a new worksheet with the title row
        const newWorksheet = xlsx.utils.json_to_sheet([formData], { header: Object.keys(formData) });
        xlsx.utils.book_append_sheet(workbook, newWorksheet, 'Form Data');
      }
    }
  
    // Save the workbook to a file
    xlsx.writeFile(workbook, 'data.xlsx');
  
    res.sendFile(__dirname + "/success.html");
  });
  
  


// Start the server
app.listen(3000, () => {
  console.log('Server is running on http://localhost:3000/');
});
