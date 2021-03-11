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

  //Leer dato de donde pegar
  await sheet.loadCells("V3:V4");
  let targetCellToPasteInfoCell = sheet.getCellByA1("V3");
  let targetCellToPaste = targetCellToPasteInfoCell.value;

  const targetCellLocation = await getCellIndexFromA1(targetCellToPaste);
  const shiftFactor = {
    row: targetCellLocation.rowIndex - allCells[0].row,
    column: targetCellLocation.columnIndex - allCells[0].column,
  };

  const budgetMatrixInfo = await getMatrixSizeFromA1(budgetLocation);
  const pasteCellRange = await createCellRangeStringFromA1AndMatrixSize(
    targetCellToPaste,
    budgetMatrixInfo,
    3
  );

  console.log(targetCellToPaste, budgetMatrixInfo, pasteCellRange);

  //Create red stripe

  const stripeSize = 19;
  let cellToPutRedStripeIndex = await getCellIndexFromA1(targetCellToPaste);
  cellToPutRedStripeIndex.rowIndex -= 2;

  const cellToPutRedStripe = await getA1StringFromIndex(
    cellToPutRedStripeIndex
  );

  const cellRangeForRedStripes = await createCellRangeStringFromA1AndMatrixSize(
    cellToPutRedStripe,
    {
      rowSize: 1,
      columnSize: stripeSize,
    }
  );

  await sheet.loadCells(cellRangeForRedStripes);
  for (
    let i = cellToPutRedStripeIndex.columnIndex;
    i < cellToPutRedStripeIndex.columnIndex + stripeSize;
    i++
  ) {
    let currCell = sheet.getCell(cellToPutRedStripeIndex.rowIndex, i);
    currCell.backgroundColor = { red: 1 };
  }
  //let currCell = sheet;

  //Write content in new cells
  await sheet.loadCells(pasteCellRange);
  allCells.forEach((cell) => {
    let currCell = sheet.getCell(
      cell.row + shiftFactor.row,
      cell.column + shiftFactor.column
    );
    currCell.value = cell.value;
  });
  await sheet.saveUpdatedCells();

  //Replace value for next paste
  const newValueToPaste = await generateNewPasteCell(pasteCellRange, 3);
  await sheet.loadCells("V3:V4");
  targetCellToPasteInfoCell = sheet.getCellByA1("V3");
  targetCellToPasteInfoCell.value = newValueToPaste;
  await sheet.saveUpdatedCells();
};

const generateNewPasteCell = async (cellRange, MARGIN = 0) => {
  let [start, end] = cellRange.split(":");
  const column = start.match(/\D+/)[0];
  let row = end.match(/\d+/)[0];
  row = parseInt(row) + MARGIN;

  const result = column + "" + row;
  return result;
};

const getCellIndexFromA1 = async (str) => {
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

async function getMatrixSizeFromA1(str) {
  let [start, end] = str.split(":");

  const startIndexes = await getCellIndexFromA1(start);
  const endIndexes = await getCellIndexFromA1(end);

  const lettersGap = endIndexes.columnIndex - startIndexes.columnIndex + 1;
  const numbersGap = endIndexes.rowIndex - startIndexes.rowIndex + 1;

  return {
    columnStartIndex: startIndexes.columnIndex,
    rowStartIndex: startIndexes.rowIndex,
    rowSize: numbersGap,
    columnSize: lettersGap,
  };
}

async function createCellRangeStringFromA1AndMatrixSize(
  cell,
  matrix,
  MARGIN = 0
) {
  const startCellIndex = await getCellIndexFromA1(cell);
  const endCellIndex = {
    rowIndex: startCellIndex.rowIndex + matrix.rowSize - 1,
    columnIndex: startCellIndex.columnIndex + matrix.columnSize - 1,
  };

  //Adding an extra marging for safety
  endCellIndex.rowIndex += MARGIN;
  endCellIndex.columnIndex += MARGIN;

  const endCellString = await getA1StringFromIndex(endCellIndex);
  const result = cell + ":" + endCellString;

  return result;
}

async function getA1StringFromIndex(cell) {
  const row = cell.rowIndex + 1;
  const column = String.fromCharCode(
    cell.columnIndex + CHAR_NUMBER_TO_SUBTRACT
  );
  const result = column + "" + row;
  return result;
}

module.exports = backupMonthlyBudget;
