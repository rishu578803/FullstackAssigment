
const multer = require('multer');
const ExcelJS = require('exceljs');
const readExcelFile = async (req, res) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);
      const worksheet = workbook.worksheets[0];
  
      let formattedData = [];
  
      worksheet.eachRow((row, rowNumber) => {
        let rowData = {};
        row.eachCell((cell, colNumber) => {
          // Check if the cell is merged
          const mergeCell = worksheet._merges[cell.address];
          if (mergeCell) {
            // Get master cell (top-left cell of the merge range)
            const masterCell = worksheet.getCell(mergeCell.top, mergeCell.left);
            rowData[`cell_${colNumber}`] = getCellStyles(masterCell, true);
          } else {
            rowData[`cell_${colNumber}`] = getCellStyles(cell, false);
          }
        });
        formattedData.push(rowData);
      });
  
      res.json(formattedData);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Failed to process the uploaded file' });
    }
  };



  const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

function getCellStyles(cell, isMerged) {
  return {
    text: cell.value,
    fontFamily: cell.font ? cell.font.name : undefined,
    fontWeight: cell.font ? (cell.font.bold ? 'bold' : 'normal') : undefined,
    textAlign: cell.alignment ? cell.alignment.horizontal : undefined,
    backgroundColor: cell.fill && cell.fill.fgColor ? cell.fill.fgColor.argb : undefined,
    textColor: cell.font && cell.font.color ? cell.font.color.argb : undefined,
    isCellMerged: isMerged
  };
}



const handleExcelFile = async (req, res) => {
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);
      const worksheet = workbook.worksheets[0];
  
      let formattedData = [];
  
      worksheet.eachRow((row, rowNumber) => {
        let rowData = {};
        row.eachCell((cell, colNumber) => {
          // Check if the cell is merged
          const mergeCell = worksheet._merges[cell.address];
          if (mergeCell) {
            // Get master cell (top-left cell of the merge range)
            const masterCell = worksheet.getCell(mergeCell.top, mergeCell.left);
            rowData[`cell_${colNumber}`] = getCellStyles(masterCell, true);
          } else {
            rowData[`cell_${colNumber}`] = getCellStyles(cell, false);
          }
        });
        formattedData.push(rowData);
      });
  
      res.json(formattedData);
    } catch (error) {
      console.error(error);
      res.status(500).json({ error: 'Failed to process the uploaded file' });
    }
  };


module.exports = {
    handleExcelFile,
    // uploadExcelFile,
};
