const express = require('express');
const excelRoute = express.Router();
const { handleExcelFile } = require('../controller/excelController');



excelRoute.post('/upload', handleExcelFile);

module.exports = {excelRoute};