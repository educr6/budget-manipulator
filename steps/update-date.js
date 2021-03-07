const { GoogleSpreadsheet } = require("google-spreadsheet");
const { DateTime } = require("luxon");

const monthsInSpanish = {
  1: "enero",
  2: "febrero",
  3: "marzo",
  4: "abril",
  5: "mayo",
  6: "junio",
  7: "julio",
  8: "agosto",
  9: "septiembre",
  10: "octubre",
  11: "noviembre",
  12: "diciembre",
};

const sheetTitles = {
  MONTHLY_BUDGET: "Presupuesto mensual",
  EXPENSE_TRACKING: "Expense tracking",
};

const updateDate = async (
  doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)
) => {
  const dateToWriteInTheSheet = await createDateString();

  //Writing date in the monthly budget sheet
  let sheet = doc.sheetsByTitle[sheetTitles.MONTHLY_BUDGET];
  await sheet.loadCells("A21:B23");
  const dateCellInMonthlyBudgetSheet = sheet.getCellByA1("A21");
  dateCellInMonthlyBudgetSheet.value = dateToWriteInTheSheet;
  await sheet.saveUpdatedCells();

  console.log("New date written in montly budget sheet");

  //Writing date in the expense tracking sheet
  sheet = doc.sheetsByTitle[sheetTitles.EXPENSE_TRACKING];
  await sheet.loadCells("A2:J2");
  const dateCellInExpenseTrackingSheet = sheet.getCellByA1("A2");
  dateCellInExpenseTrackingSheet.value = dateToWriteInTheSheet;
  await sheet.saveUpdatedCells();

  console.log("New date written in expense tracking sheet");
};

const createDateString = async () => {
  const todaysDate = DateTime.now();
  const currentMonth = monthsInSpanish[todaysDate.month];

  const nextMonth =
    currentMonth == monthsInSpanish[12]
      ? monthsInSpanish[0]
      : monthsInSpanish[todaysDate.month + 1];

  let dateToWriteInTheSheet = `${todaysDate.day} de ${currentMonth} a 27 de ${nextMonth}`;
  return dateToWriteInTheSheet;
};

module.exports = updateDate;
