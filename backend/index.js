
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const app = express();
require('dotenv').config();


const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const worksheet = workbook.getWorksheet(1); 

    const result = [];

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        const cellValue = cell.value;
        const cellStyle = cell.style;

        const textColor = cellStyle.font?.color?.argb;
        const fontFamily = cellStyle.font?.name;
        const fontWeight = cellStyle.font?.bold ? 'bold' : 'normal';
        const textAlign = cellStyle.alignment?.horizontal;
        const backgroundColor = cellStyle.fill?.fgColor?.argb;

        result.push({
          cell: `${rowNumber}:${colNumber}`,
          value: cellValue,
          textColor: textColor,
          fontFamily: fontFamily,
          fontWeight: fontWeight,
          textAlign: textAlign,
          backgroundColor: backgroundColor,
        });
      });
    });

    res.json(result);
  } catch (error) {
    res.status(500).send(error.message);
  }
});


const port = process.env.PORT || 7000;
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
