const { GoogleSpreadsheet } = require("google-spreadsheet");
const CHAR_BASE_NUMBER = 65;
const months_in_spanish = require("../../months_in_spanish");
const { DateTime } = require("luxon");

const backupData = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID), dataToBackup) => {
  const dataToBeBackedUp = await copyData(doc, dataToBackup);
  dataToBackup.dataToBeBackedUp = dataToBeBackedUp;
  console.log("The data copied was ", dataToBeBackedUp);

  const cellWhereBackupWillBePasted = await getCellWhereBackupWillBePasted(doc, dataToBackup);
  dataToBackup.cellWhereBackupWillBePasted = cellWhereBackupWillBePasted;
  console.log("Cell where backup will be pasted is ", cellWhereBackupWillBePasted);

  const shiftFactor = await calculateShiftFactor(dataToBeBackedUp, cellWhereBackupWillBePasted);
  dataToBackup.shiftFactor = shiftFactor;
  console.log("Shiftfactor is  ", shiftFactor);

  const dataToBeBackedUpMatrixInfo = await getMatrixSizeFromA1(dataToBackup.dataLocation);
  dataToBackup.dataToBeBackedUpMatrixInfo = dataToBeBackedUpMatrixInfo;
  console.log("Matrix info  ", dataToBeBackedUpMatrixInfo);

  const pasteMargin = 3;
  const pasteCellRange = await createCellRangeStringFromA1AndMatrixSize(
    cellWhereBackupWillBePasted,
    dataToBeBackedUpMatrixInfo,
    pasteMargin
  );
  dataToBackup.pasteCellRange = pasteCellRange;
  console.log("PasteCell range is  ", pasteCellRange);

  await createRedStripeWithMonth(doc, dataToBackup);
  await writeContentOfBackup(doc, dataToBackup);
  await replacePasteLocationForFutureBackups(doc, dataToBackup);
};

const replacePasteLocationForFutureBackups = async (doc, dataToBackup) => {
  const MARGIN_FOR_FUTURE_BACKUP = 3;
  const newValueToPasteBackups = await generateNewPasteCell(dataToBackup.pasteCellRange, MARGIN_FOR_FUTURE_BACKUP);
  let sheet = doc.sheetsByTitle[dataToBackup.pasteSheet];
  await sheet.loadCells(dataToBackup.pasteCellRange);

  let targetCellToPasteInfoCell = sheet.getCellByA1(dataToBackup.pasteCell);
  targetCellToPasteInfoCell.value = newValueToPasteBackups;
  await sheet.saveUpdatedCells();
};

const writeContentOfBackup = async (doc, dataToBackup) => {
  let sheet = doc.sheetsByTitle[dataToBackup.pasteSheet];
  await sheet.loadCells(dataToBackup.pasteCellRange);

  dataToBackup.dataToBeBackedUp.forEach((cell) => {
    let currCell = sheet.getCell(
      cell.row + dataToBackup.shiftFactor.row,
      cell.column + dataToBackup.shiftFactor.column
    );
    currCell.value = cell.value;
  });

  await sheet.saveUpdatedCells();
};

const createRedStripeWithMonth = async (doc, dataToBackup) => {
  const STRIPE_SIZE = 19;
  let cellToPutRedStripeIndex = await getCellIndexFromA1(dataToBackup.cellWhereBackupWillBePasted);
  cellToPutRedStripeIndex.rowIndex -= 2;
  const cellToPutRedStripe = await getA1StringFromIndex(cellToPutRedStripeIndex);
  const matrixWithOneRowAndXColumns = { rowSize: 1, columnSize: STRIPE_SIZE };
  const cellRangeForRedStripes = await createCellRangeStringFromA1AndMatrixSize(
    cellToPutRedStripe,
    matrixWithOneRowAndXColumns
  );

  //WRITE MONTH IN THE STRIPE
  let sheet = doc.sheetsByTitle[dataToBackup.pasteSheet];
  await sheet.loadCells(cellRangeForRedStripes);

  let cellToWriteMonth = sheet.getCellByA1(cellToPutRedStripe);
  cellToWriteMonth.textFormat = { bold: true, fontSize: 24 };
  const todaysDate = DateTime.now();
  const currentMonth = monthsInSpanish[todaysDate.month];
  cellToWriteMonth.value = "" + currentMonth + " " + todaysDate.year;

  //COLOR THE STRIPE RED
  for (let i = cellToPutRedStripeIndex.columnIndex; i < cellToPutRedStripeIndex.columnIndex + STRIPE_SIZE; i++) {
    let currCell = sheet.getCell(cellToPutRedStripeIndex.rowIndex, i);
    currCell.backgroundColor = { red: 1 };
  }

  await sheet.saveUpdatedCells();
};

const calculateShiftFactor = async (dataToBeBackedUp, cellWhereBackupWillBePasted) => {
  const targetCellLocation = await getCellIndexFromA1(cellWhereBackupWillBePasted);
  const shiftFactor = {
    row: targetCellLocation.rowIndex - dataToBeBackedUp[0].row,
    column: targetCellLocation.columnIndex - dataToBeBackedUp[0].column,
  };

  return shiftFactor;
};

const copyData = async (doc, dataToBackup) => {
  let sheet = doc.sheetsByTitle[dataToBackup.dataSheet];
  await sheet.loadCells(dataToBackup.dataLocation);
  let copiedData = [];
  const matrix = await getMatrixSizeFromA1(dataToBackup.dataLocation);

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    for (let j = matrix.columnStartIndex; j < matrix.columnStartIndex + matrix.columnSize; j++) {
      let currentCell = sheet.getCell(i, j);
      copiedData.push({ column: j, row: i, value: currentCell.value });
    }
  }

  return copiedData;
};

const getCellWhereBackupWillBePasted = async (doc, dataToBackup) => {
  let sheet = doc.sheetsByTitle[dataToBackup.pasteSheet];
  await sheet.loadCells(dataToBackup.pasteCellRange);
  let targetCellToPasteInfoCell = sheet.getCellByA1(dataToBackup.pasteCell);
  let targetCellToPaste = targetCellToPasteInfoCell.value;

  return targetCellToPaste;
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
  const CHAR_BASE_NUMBER = 65;

  const matchNumbersRegex = /\d+/;
  const matchNonNumbersRegex = /\D+/;

  let startNumber = str.match(matchNumbersRegex);
  startNumber = parseInt(startNumber[0], "10");

  let startLetter = str.match(matchNonNumbersRegex);
  startLetter = startLetter[0];

  const rowIndex = startNumber - 1;
  const columnIndex = startLetter.charCodeAt(0) - CHAR_BASE_NUMBER;

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

async function createCellRangeStringFromA1AndMatrixSize(cell, matrix, MARGIN = 0) {
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
  const column = String.fromCharCode(cell.columnIndex + CHAR_BASE_NUMBER);
  const result = column + "" + row;
  return result;
}

module.exports = backupData;
