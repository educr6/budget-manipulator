require("dotenv").config();
const creds = require("./budget-manipulator-d57be5158174.json");

(async function () {
  const { GoogleSpreadsheet } = require("google-spreadsheet");

  // Initialize the sheet - doc ID is the long id in the sheets URL
  const doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID);

  // Initialize Auth - see more available options at https://theoephraim.github.io/node-google-spreadsheet/#/getting-started/authentication
  await doc.useServiceAccountAuth(creds);

  await doc.loadInfo(); // loads document properties and worksheets
  console.log(doc.title);

  const sheet = doc.sheetsByIndex[0]; // or use doc.sheetsById[id] or doc.sheetsByTitle[title]
  console.log(sheet.title);
  console.log(sheet.rowCount);

  // adding / removing sheets
  const newSheet = await doc.addSheet({ title: "hot new sheet!" });
  await newSheet.delete();
})();
