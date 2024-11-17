const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { processExcelSheets } = require('./excelProcessor');

const app = express();
const port = process.env.PORT || 3000;

// Set up EJS as the view engine
app.set('view engine', 'ejs');

// Serve static files
app.use(express.static('public'));

// Configure multer for file uploads
const upload = multer({ dest: 'uploads/' });

// Create uploads directory if it doesn't exist
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Routes
app.get('/', (req, res) => {
  res.render('index');
});

app.post('/upload', upload.fields([
  { name: 'mainSheet', maxCount: 1 },
  { name: 'bmSheet', maxCount: 1 },
  { name: 'dmmSheet', maxCount: 1 },
  { name: 'hoSheet', maxCount: 1 },
  { name: 'hooplaSheet', maxCount: 1 }
]), async (req, res) => {
  try {
    if (!req.files.mainSheet) {
      throw new Error('Main sheet is required');
    }

    const mainSheetPath = req.files.mainSheet[0].path;
    const otherSheetPaths = [];

    // Collect paths of uploaded reference sheets
    ['bmSheet', 'dmmSheet', 'hoSheet', 'hooplaSheet'].forEach(sheetName => {
      if (req.files[sheetName]) {
        otherSheetPaths.push(req.files[sheetName][0].path);
      }
    });

    const result = await processExcelSheets(mainSheetPath, otherSheetPaths);

    // Clean up uploaded files
    fs.unlinkSync(mainSheetPath);
    otherSheetPaths.forEach(path => fs.unlinkSync(path));

    res.json({
      success: true,
      downloadUrl: `/download/${path.basename(result.outputPath)}`,
      stats: {
        totalRows: result.totalRows,
        removedRows: result.removedRows,
        remainingRows: result.remainingRows
      }
    });
  } catch (error) {
    res.status(400).json({ success: false, error: error.message });
  }
});

app.get('/download/:filename', (req, res) => {
  const filePath = path.join(__dirname, 'downloads', req.params.filename);
  res.download(filePath, 'filtered_main_sheet.xlsx', err => {
    if (!err) {
      // Delete the file after download
      fs.unlinkSync(filePath);
    }
  });
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});