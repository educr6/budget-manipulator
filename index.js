require("dotenv").config();
const creds = require(process.env.GOOGLE_CREDENTIALS_ROUTE);
const updateDate = require("./steps/update-date");
const backupMonthlyBudget = require("./steps/backup/monthly-budget");
const backupExpenseTracking = require("./steps/backup/expense-tracking");

(async function () {
  const { GoogleSpreadsheet } = require("google-spreadsheet");

  // Initialize the sheet - doc ID is the long id in the sheets URL
  const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID);

  // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
  await doc.useServiceAccountAuth(creds);
  await doc.loadInfo(); // loads document properties and worksheets
  console.log(doc.title);

  //updateDate(doc);
  //await backupMonthlyBudget(doc);
  await backupExpenseTracking(doc);
})();
