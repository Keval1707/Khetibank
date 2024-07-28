const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Initialize Express app
const app = express();
const PORT = process.env.PORT || 5000;

// Serve static files
app.use(express.static('public'));

// Configure multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Set EJS as the view engine
app.set('view engine', 'ejs');

// Convert Excel serial date to JavaScript Date object
const excelDateToJSDate = (serial) => {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  return new Date(utc_value * 1000);
};

// Helper function to count occurrences
const countOccurrences = (arr) => {
  return arr.reduce((acc, val) => {
    if (val) {
      acc[val] = (acc[val] || 0) + 1;
    }
    return acc;
  }, {});
};

// Route to render the upload form
app.get('/', (req, res) => {
  res.render('upload');
});

// Route to handle file upload and processing
app.post('/upload', upload.single('file'), (req, res) => {
  const filePath = req.file.path;

  // Read the Excel file
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  // Count occurrences of 'Recruiter Code', skipping NaN values
  const recruiterCodes = sheet.map(row => row['Recruiter Code']);
  let recruiterCodeCounts = countOccurrences(recruiterCodes);

  // Filter to include only counts greater than 1
  recruiterCodeCounts = Object.entries(recruiterCodeCounts)
    .filter(([key, count]) => count > 1)
    .reduce((acc, [key, count]) => {
      acc[key] = count;
      return acc;
    }, {});

  // Count occurrences of 'Application Code', skipping NaN values
  const applicationCodes = sheet.map(row => row['Application Code']);
  let applicationCodeCounts = countOccurrences(applicationCodes);

  // Calculate total Application Code count
  const totalApplicationCodeCount = Object.values(applicationCodeCounts).reduce((acc, count) => acc + count, 0);

  // Count applications received by date
  const applicationsReceivedDates = sheet.map(row => {
    const dateValue = row['Applications Received Date'];
    if (typeof dateValue === 'number') {
      return excelDateToJSDate(dateValue).toISOString().split('T')[0];
    }
    return dateValue;
  });
  const applicationsReceivedByDate = countOccurrences(applicationsReceivedDates);

  // Prepare data for new Excel file
  const outputWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputWorkbook, XLSX.utils.json_to_sheet(Object.entries(recruiterCodeCounts).map(([key, count]) => ({ 'Recruiter Code': key, 'Count': count }))), 'Recruiter Code Counts');

  let applicationCodeCountsArray = Object.entries(applicationCodeCounts).map(([key, count]) => ({ 'Application Code': key, 'Count': count }));
  applicationCodeCountsArray.push({ 'Application Code': 'Total', 'Count': totalApplicationCodeCount });
  XLSX.utils.book_append_sheet(outputWorkbook, XLSX.utils.json_to_sheet(applicationCodeCountsArray), 'Application Code Counts');

  XLSX.utils.book_append_sheet(outputWorkbook, XLSX.utils.json_to_sheet(Object.entries(applicationsReceivedByDate).map(([key, count]) => ({ 'Applications Received Date': key, 'Count': count }))), 'Applications Received by Date');

  // Save counts to a new Excel file
  const outputFilePath = path.join(__dirname, 'public', 'counts.xlsx');
  XLSX.writeFile(outputWorkbook, outputFilePath);

  // Send response
  res.render('result', {
    recruiterCodeCounts,
    applicationCodeCounts,
    totalApplicationCodeCount,
    applicationsReceivedByDate,
    outputFilePath: '/counts.xlsx'
  });

  // Clean up uploaded file
  fs.unlinkSync(filePath);
});

// Route to handle file download
app.get('/download', (req, res) => {
  const file = path.join(__dirname, 'public', 'counts.xlsx');
  res.download(file);
});

// Start the server
app.listen(PORT, () => {
  console.log(`server IP:  http://localhost:${PORT}\n`);
});
