const { GoogleSpreadsheet } = require("google-spreadsheet");

const backupMonthlyBudget = async (
  doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)
) => {
  console.log("Hello world");

  let sheet = doc.sheetsByTitle["backup-test"];
  await sheet.loadCells("A3:B5");
  let allCells = [];
  const matrix = await getMatrixSizeFromA1("A3:B5");

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

  let targetCellToPaste = "E3";
  const targetCellLocation = getCellIndexFromA1(targetCellToPaste);
  const shiftFactor = {
    row: targetCellLocation.rowIndex - allCells[0].row,
    column: targetCellLocation.columnIndex - allCells[0].column,
  };

  await sheet.loadCells("A1:S90");
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
  const CHAR_NUMBER_TO_SUBTRACT = 65;

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
  const matchNumbersRegex = /\d+/;
  const matchNonNumbersRegex = /\D+/;

  let startNumber = start.match(matchNumbersRegex);
  let endNumber = end.match(matchNumbersRegex);

  startNumber = parseInt(startNumber[0], "10");
  endNumber = parseInt(endNumber, "10");

  let startLetter = start.match(matchNonNumbersRegex);
  let endLetter = end.match(matchNonNumbersRegex);

  startLetter = startLetter[0];
  endLetter = endLetter[0];

  const numbersGap = endNumber - startNumber + 1;
  const lettersGap = endLetter.charCodeAt(0) - startLetter.charCodeAt(0) + 1;

  const columnStart = startLetter.charCodeAt(0) - CHAR_NUMBER_TO_SUBTRACT;

  return {
    columnStartIndex: columnStart,
    rowStartIndex: startNumber - 1,
    rowSize: numbersGap,
    columnSize: lettersGap,
  };
}

module.exports = backupMonthlyBudget;
