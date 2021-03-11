const { GoogleSpreadsheet } = require("google-spreadsheet");

const sheetTitles = {
  MONTHLY_BUDGET: "Presupuesto mensual",
  EXPENSE_TRACKING: "Expense tracking",
};

const CHAR_NUMBER_TO_SUBTRACT = 65;

const backupMonthlyBudget = async (
  doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)
) => {
  const budgetLocation = "A2:S29";
  let sheet = doc.sheetsByTitle[sheetTitles.MONTHLY_BUDGET];
  await sheet.loadCells(budgetLocation);
  let allCells = [];
  const matrix = await getMatrixSizeFromA1(budgetLocation);

  for (
    let i = matrix.rowStartIndex;
    i < matrix.rowStartIndex + matrix.rowSize;
    i++
  ) {
    for (
      let j = matrix.columnStartIndex;
      j < matrix.columnStartIndex + matrix.columnSize;
      j++
    ) {
      let currentCell = sheet.getCell(i, j);
      console.log(
        "This is the value of " + i + "," + j + ": " + currentCell.value
      );

      allCells.push({
        column: j,
        row: i,
        value: currentCell.value,
      });
    }
  }

  let targetCellToPaste = "A37";
  const targetCellLocation = getCellIndexFromA1(targetCellToPaste);
  const shiftFactor = {
    row: targetCellLocation.rowIndex - allCells[0].row,
    column: targetCellLocation.columnIndex - allCells[0].column,
  };

  await sheet.loadCells("A37:S69");
  allCells.forEach((cell) => {
    let currCell = sheet.getCell(
      cell.row + shiftFactor.row,
      cell.column + shiftFactor.column
    );
    currCell.value = cell.value;
  });
  await sheet.saveUpdatedCells();
};

const getCellIndexFromA1 = (str) => {
  const matchNumbersRegex = /\d+/;
  const matchNonNumbersRegex = /\D+/;

  let startNumber = str.match(matchNumbersRegex);
  startNumber = parseInt(startNumber[0], "10");

  let startLetter = str.match(matchNonNumbersRegex);
  startLetter = startLetter[0];

  const rowIndex = startNumber - 1;
  const columnIndex = startLetter.charCodeAt(0) - CHAR_NUMBER_TO_SUBTRACT;

  return {
    rowIndex: rowIndex,
    columnIndex: columnIndex,
  };
};

function getMatrixSizeFromA1(str) {
  const CHAR_NUMBER_TO_SUBTRACT = 65;
  let [start, end] = str.split(":");

  const startIndexes = getCellIndexFromA1(start);
  const endIndexes = getCellIndexFromA1(end);

  const lettersGap = endIndexes.columnIndex - startIndexes.columnIndex + 1;
  const numbersGap = endIndexes.rowIndex - startIndexes.rowIndex + 1;

  return {
    columnStartIndex: startIndexes.columnIndex,
    rowStartIndex: startIndexes.rowIndex,
    rowSize: numbersGap,
    columnSize: lettersGap,
  };
}

module.exports = backupMonthlyBudget;
