const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { processExcelSheets } = require('./excelProcessor');

const app = express();
const port = process.env.PORT || 3000; // Updated for Railway

// Rest of your server.js code remains exactly the same