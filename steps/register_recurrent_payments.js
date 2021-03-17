const { GoogleSpreadsheet } = require("google-spreadsheet");
const sheetTitles = require("../sheet_titles");
const { getMatrixSizeFromA1, getA1StringFromIndex, getCellIndexFromA1 } = require("./backup/base-methods");

const expensesCellRange = "A5:J73";

const isCellEmpty = (cell) => {
  return cell.value === null;
};

const registerRecurrentPayments = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)) => {
  let sheet = doc.sheetsByTitle[sheetTitles.EXPENSE_TRACKING];
  await sheet.loadCells(expensesCellRange);
  const matrix = await getMatrixSizeFromA1(expensesCellRange);

  console.log("This is the matrix created from " + expensesCellRange + ": ", matrix);

  let cellWhereNewExpenseWillBeRegister = "";

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    let currentCell = sheet.getCell(i, 0);
    if (isCellEmpty(currentCell)) {
      const cell = await getA1StringFromIndex({ rowIndex: i, columnIndex: 0 });
      console.log("The answer is " + cell);
      cellWhereNewExpenseWillBeRegister = cell;
      break;
    } else {
      console.log(currentCell.value);
    }
  }

  const randomExpense = {
    type: "Gastos fijos",
    description: "Written from javascript",
    category: "Otros",
    amount: 1300,
    paymentMethod: "credito",
  };

  let descriptionCellIndex = await getCellIndexFromA1(cellWhereNewExpenseWillBeRegister);
  let categoryCellIndex = JSON.parse(JSON.stringify(descriptionCellIndex));
  categoryCellIndex.columnIndex += 1;
  let amountCellIndex = JSON.parse(JSON.stringify(descriptionCellIndex));
  amountCellIndex.columnIndex += 2;
  let paymentMethodCellIndex = JSON.parse(JSON.stringify(descriptionCellIndex));
  paymentMethodCellIndex.columnIndex += 3;

  let descriptionCell = sheet.getCell(descriptionCellIndex.rowIndex, descriptionCellIndex.columnIndex);
  descriptionCell.value = randomExpense.description;

  let categoryCell = sheet.getCell(categoryCellIndex.rowIndex, categoryCellIndex.columnIndex);
  categoryCell.value = randomExpense.category;

  let amountCell = sheet.getCell(amountCellIndex.rowIndex, amountCellIndex.columnIndex);
  amountCell.value = randomExpense.amount;

  let paymentMethodCell = sheet.getCell(paymentMethodCellIndex.rowIndex, paymentMethodCellIndex.columnIndex);
  paymentMethodCell.value = randomExpense.paymentMethod;
  await sheet.saveUpdatedCells();
};

module.exports = registerRecurrentPayments;
