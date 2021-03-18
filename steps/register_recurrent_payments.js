const { GoogleSpreadsheet } = require("google-spreadsheet");
const sheetTitles = require("../sheet_titles");
const { getMatrixSizeFromA1, getA1StringFromIndex, getCellIndexFromA1 } = require("./backup/base-methods");

const expensesCellRange = "A5:J73";

const isCellEmpty = (cell) => {
  return cell.value === null;
};

const expenseTypes = {
  LIVING_EXPENSE: "Gastos fijos",
  FINANCIAL_IMPROVEMENT: "Mejoria financiera",
  FREE_SPENDING: "Gastos discresionales",
};

const registerRecurrentPayments = async (doc = new GoogleSpreadsheet(process.env.GOOGLE_SHEET_ID)) => {
  let sheet = doc.sheetsByTitle[sheetTitles.EXPENSE_TRACKING];
  await sheet.loadCells(expensesCellRange);
  const matrix = await getMatrixSizeFromA1(expensesCellRange);

  console.log("This is the matrix created from " + expensesCellRange + ": ", matrix);

  let sheetInfo = {
    sheet: sheet,
    matrix: matrix,
  };

  const automaticExpenses = [
    {
      type: expenseTypes.LIVING_EXPENSE,
      description: "Seguro de mami, from js",
      category: "Seguro mami",
      amount: 1510,
      paymentMethod: "debito",
    },
    {
      type: expenseTypes.FINANCIAL_IMPROVEMENT,
      description: "Deduccion prestamo, from js",
      amount: 8720,
    },
  ];

  automaticExpenses.forEach(async (expense) => {
    await registerExpense(expense, sheetInfo);
  });
};

const locateCellToRegisterExpense = async (expense, sheetInfo) => {
  let matrix = sheetInfo.matrix;
  let sheet = sheetInfo.sheet;

  const colum = await getColumnIndexBasedOnExpenseType(expense);

  let cellWhereNewExpenseWillBeRegistered = "";

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    let currentCell = sheet.getCell(i, colum);
    console.log("Value of this row is ", currentCell.value);
    console.log("Cell being checked is ", await getA1StringFromIndex({ rowIndex: i, columnIndex: colum }));

    if (isCellEmpty(currentCell)) {
      const cell = await getA1StringFromIndex({ rowIndex: i, columnIndex: colum });
      cellWhereNewExpenseWillBeRegistered = cell;
      break;
    }
  }

  return cellWhereNewExpenseWillBeRegistered;
};

const getColumnIndexBasedOnExpenseType = async (expense) => {
  let result = 0;

  switch (expense.type) {
    case expenseTypes.LIVING_EXPENSE:
      result = 0;
      break;
    case expenseTypes.FREE_SPENDING:
      result = 4;
      break;
    case expenseTypes.FINANCIAL_IMPROVEMENT:
      result = 8;
      break;
    default:
      throw "Expense type not recognized";
  }

  return result;
};

const getAmountColumnShiftFactor = async (expense) => {
  if (expense.type === expenseTypes.FINANCIAL_IMPROVEMENT) {
    return 1;
  }

  return 2;
};

const getCategoryColumnShiftFactor = async (expense) => {
  return 1;
};

const getPaymentMethodColumnShiftFactor = async (expense) => {
  return 3;
};

const copyObject = async (obj) => {
  return JSON.parse(JSON.stringify(obj));
};

const registerExpense = async (expense, sheetInfo) => {
  let sheet = sheetInfo.sheet;
  let cellWhereExpenseWillBeRegistered = await locateCellToRegisterExpense(expense, sheetInfo);
  console.log("La celda de la discordia: ", cellWhereExpenseWillBeRegistered);
  let descriptionCellIndex = await getCellIndexFromA1(cellWhereExpenseWillBeRegistered);

  let amountCellIndex = await copyObject(descriptionCellIndex);
  amountCellIndex.columnIndex += await getAmountColumnShiftFactor(expense);
  console.log("La de amount: ", amountCellIndex);

  let descriptionCell = sheet.getCell(descriptionCellIndex.rowIndex, descriptionCellIndex.columnIndex);
  descriptionCell.value = expense.description;

  let amountCell = sheet.getCell(amountCellIndex.rowIndex, amountCellIndex.columnIndex);
  amountCell.value = expense.amount;

  if (expense.type !== expenseTypes.FINANCIAL_IMPROVEMENT) {
    let categoryCellIndex = await copyObject(descriptionCellIndex);
    categoryCellIndex.columnIndex += await getCategoryColumnShiftFactor(expense);

    let paymentMethodCellIndex = await copyObject(descriptionCellIndex);
    paymentMethodCellIndex.columnIndex += await getPaymentMethodColumnShiftFactor(expense);

    let categoryCell = sheet.getCell(categoryCellIndex.rowIndex, categoryCellIndex.columnIndex);
    categoryCell.value = expense.category;

    let paymentMethodCell = sheet.getCell(paymentMethodCellIndex.rowIndex, paymentMethodCellIndex.columnIndex);
    paymentMethodCell.value = expense.paymentMethod;
  }

  await sheet.saveUpdatedCells();
};

module.exports = registerRecurrentPayments;
