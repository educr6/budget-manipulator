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

  let sheetInfo = {
    sheet: sheet,
    matrix: matrix,
  };

  const automaticExpenses = [
    {
      type: expenseTypes.LIVING_EXPENSE,
      description: "Gastos comunes supermercado, from js",
      category: "Supermercado",
      formula: "='Gastos super'!$D$11",
      paymentMethod: "debito",
    },
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

  let cellsToRegisterExpense = await locateCellToRegisterExpense(sheetInfo);

  for (let i = 0; i < automaticExpenses.length; i++) {
    cellsToRegisterExpense = await registerExpense(automaticExpenses[i], sheetInfo, cellsToRegisterExpense);
  }
};

const locateCellToRegisterExpense = async (sheetInfo) => {
  let matrix = sheetInfo.matrix;
  let sheet = sheetInfo.sheet;

  const livingExpenseColumn = 0;
  const freeSpendingColum = 4;
  const financialImprovementColumn = 8;

  let cellWhereLivingExpenseWillBeRegistered = "";
  let cellWhereFreeExpenseWillBeRegistered = "";
  let cellWhereFinancialImprovementWillBeRegistered = "";

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    let currentCell = sheet.getCell(i, livingExpenseColumn);
    if (isCellEmpty(currentCell)) {
      let cell = await getA1StringFromIndex({ rowIndex: i, columnIndex: livingExpenseColumn });
      cellWhereLivingExpenseWillBeRegistered = cell;
      break;
    }
  }

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    let currentCell = sheet.getCell(i, freeSpendingColum);
    if (isCellEmpty(currentCell)) {
      let cell = await getA1StringFromIndex({ rowIndex: i, columnIndex: freeSpendingColum });
      cellWhereFreeExpenseWillBeRegistered = cell;
      break;
    }
  }

  for (let i = matrix.rowStartIndex; i < matrix.rowStartIndex + matrix.rowSize; i++) {
    let currentCell = sheet.getCell(i, financialImprovementColumn);
    if (isCellEmpty(currentCell)) {
      let cell = await getA1StringFromIndex({ rowIndex: i, columnIndex: financialImprovementColumn });
      cellWhereFinancialImprovementWillBeRegistered = cell;
      break;
    }
  }

  let result = {};
  result[expenseTypes.LIVING_EXPENSE] = cellWhereLivingExpenseWillBeRegistered;
  result[expenseTypes.FREE_SPENDING] = cellWhereFreeExpenseWillBeRegistered;
  result[expenseTypes.FINANCIAL_IMPROVEMENT] = cellWhereFinancialImprovementWillBeRegistered;
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

const registerExpense = async (expense, sheetInfo, cellsToRegisterExpenses) => {
  let sheet = sheetInfo.sheet;
  let cellWhereExpenseWillBeRegistered = cellsToRegisterExpenses[expense.type];
  console.log("Expense being written: ", expense.description);

  let descriptionCellIndex = await getCellIndexFromA1(cellWhereExpenseWillBeRegistered);

  let amountCellIndex = await copyObject(descriptionCellIndex);
  amountCellIndex.columnIndex += await getAmountColumnShiftFactor(expense);

  let descriptionCell = sheet.getCell(descriptionCellIndex.rowIndex, descriptionCellIndex.columnIndex);
  descriptionCell.value = expense.description;

  let amountCell = sheet.getCell(amountCellIndex.rowIndex, amountCellIndex.columnIndex);
  amountCell.value = expense.amount;
  if (expense.hasOwnProperty("formula")) {
    amountCell.formula = expense.formula;
  }

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

  let newCellIndex = await copyObject(descriptionCellIndex);
  newCellIndex.rowIndex += 1;
  cellsToRegisterExpenses[expense.type] = await getA1StringFromIndex(newCellIndex);

  await sheet.saveUpdatedCells();

  return cellsToRegisterExpenses;
};

module.exports = registerRecurrentPayments;
