const { GoogleSpreadsheet } = require("google-spreadsheet");
const sheetTitles = require("../sheet_titles");
const { getMatrixSizeFromA1 } = require("./backup/base-methods");

const expensesCellRange = "A5:J73";

const clearExpenses = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)) => {
  let sheet = doc.sheetsByTitle[sheetTitles.EXPENSE_TRACKING];
  await sheet.loadCells(expensesCellRange);
  const matrix = await getMatrixSizeFromA1(expensesCellRange);

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    for (let j = matrix.columnStartIndex; j < matrix.columnStartIndex + matrix.columnSize; j++) {
      let currentCell = sheet.getCell(i, j);
      currentCell.value = "";
    }
  }

  sheet.saveUpdatedCells();
};

module.exports = clearExpenses;
