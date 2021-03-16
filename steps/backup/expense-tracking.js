const { GoogleSpreadsheet } = require("google-spreadsheet");
const { backupData } = require("./base-methods");
const sheetTitles = require("./../../sheet_titles");

const backUpExpenseTracking = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)) => {
  const backUpInfo = {
    dataSheet: sheetTitles.EXPENSE_TRACKING,
    dataLocation: "A2:W47",
    pasteSheet: sheetTitles.EXPENSE_HISTORY,
    pasteCellRange: "T4:T5",
    pasteCell: "T4",
  };

  await backupData(doc, backUpInfo);
};

module.exports = backUpExpenseTracking;
