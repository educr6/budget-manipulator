const { GoogleSpreadsheet } = require("google-spreadsheet");
const { backupData } = require("./base-methods");
const sheetTitles = require("./../../sheet_titles");

const backUpMonthlyBudget = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)) => {
  const backUpInfo = {
    dataSheet: sheetTitles.MONTHLY_BUDGET,
    dataLocation: "A2:S29",
    pasteSheet: sheetTitles.BUDGET_HISTORY,
    pasteCellRange: "W5:W6",
    pasteCell: "W5",
  };

  await backupData(doc, backUpInfo);
};

module.exports = backUpMonthlyBudget;
