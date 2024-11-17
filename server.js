const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { processExcelSheets } = require('./excelProcessor');

const app = express();
const port = process.env.PORT || 3000;

// Set up EJS as the view engine
app.set('view engine', 'ejs');

// Create required directories
['uploads', 'downloads'].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir);
  }
});

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ storage: storage });

// Serve static files
app.use(express.static('public'));
app.use('/downloads', express.static('downloads'));

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
    if (!req.files || !req.files.mainSheet) {
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
    setTimeout(() => {
      try {
        fs.unlinkSync(mainSheetPath);
        otherSheetPaths.forEach(path => {
          if (fs.existsSync(path)) {
            fs.unlinkSync(path);
          }
        });
      } catch (err) {
        console.error('Cleanup error:', err);
      }
    }, 1000);

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
    console.error('Upload error:', error);
    res.status(400).json({ success: false, error: error.message });
  }
});

app.get('/download/:filename', (req, res) => {
  const filePath = path.join(__dirname, 'downloads', req.params.filename);
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('File not found');
  }
  
  res.download(filePath, 'filtered_main_sheet.xlsx', err => {
    if (!err) {
      // Delete the file after download
      setTimeout(() => {
        try {
          if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
          }
        } catch (err) {
          console.error('Error deleting file:', err);
        }
      }, 1000);
    }
  });
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({ success: false, error: 'Something went wrong!' });
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});