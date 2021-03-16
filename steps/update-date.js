const { GoogleSpreadsheet } = require("google-spreadsheet");
const { DateTime } = require("luxon");
const monthsInSpanish = require("../months_in_spanish");

const dateCells = {
  monthly_budget: "A21",
  expense_tracking: "A2",
};

const sheetTitles = {
  MONTHLY_BUDGET: "Presupuesto mensual",
  EXPENSE_TRACKING: "Expense tracking",
};

const updateDate = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)) => {
  const dateToWriteInTheSheet = await createDateString();

  //Writing date in the monthly budget sheet
  let sheet = doc.sheetsByTitle[sheetTitles.MONTHLY_BUDGET];

  writeDataInCell({
    cell: dateCells.monthly_budget,
    sheet: sheet,
    data: dateToWriteInTheSheet,
  });

  console.log("New date written in monthly budget sheet");

  //Writing date in the expense tracking sheet

  sheet = doc.sheetsByTitle[sheetTitles.EXPENSE_TRACKING];
  writeDataInCell({
    cell: dateCells.expense_tracking,
    sheet: sheet,
    data: dateToWriteInTheSheet,
  });

  console.log("New date written in expense tracking sheet");
};

const writeDataInCell = async (dataObject) => {
  if (!"data" in dataObject) throw "The object does not contain data to be written in the cell";
  if (!"cell" in dataObject) throw "The object does not contain the cell to write on";
  if (!"sheet" in dataObject) throw "The object does not contain the sheet object";

  let sheet = dataObject.sheet;
  await sheet.loadCells(dataObject.cell);
  const cellToWriteOn = sheet.getCellByA1(dataObject.cell);
  cellToWriteOn.value = dataObject.data;
  await sheet.saveUpdatedCells();
};

const createDateString = async () => {
  const todaysDate = DateTime.now();
  const currentMonth = monthsInSpanish[todaysDate.month];

  const nextMonth = currentMonth == monthsInSpanish[12] ? monthsInSpanish[0] : monthsInSpanish[todaysDate.month + 1];

  let dateToWriteInTheSheet = `${todaysDate.day} de ${currentMonth} a 27 de ${nextMonth}`;
  return dateToWriteInTheSheet;
};

module.exports = updateDate;
