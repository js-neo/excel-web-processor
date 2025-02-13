import { getColumnLetter } from "../utils/excelUtils.js";

const formulas = {
    totalCost: (row) => `D${row} * E${row}`,
    avrCost: (row, insertIndex) =>
        `${getColumnLetter(insertIndex)}${row} * E${row}`,
    totalQuantity: (row, calculateTotalQuantity) =>
        `SUM(${calculateTotalQuantity(row)})`,
    completedCost: (row, insertIndex) =>
        `${getColumnLetter(insertIndex + 2)}${row} * E${row}`,
    quantityRemaining: (row, insertIndex) =>
        `D${row} - ${getColumnLetter(insertIndex + 2)}${row}`,
    remainingCost: (row, insertIndex) =>
        `${getColumnLetter(insertIndex + 4)}${row} * E${row}`,
    excess: (row, prevToLastColLetter) =>
        `IF(${prevToLastColLetter}${row}<0, ABS(${prevToLastColLetter}${row}), 0)`,
    sum: (row, insertIndex, offset) =>
        `SUM(${getColumnLetter(insertIndex + offset)}2:${getColumnLetter(insertIndex + offset)}${row - 1})`,
    sumIf: (row, insertIndex, offset) =>
        `SUMIF(${getColumnLetter(insertIndex + offset)}2:${getColumnLetter(insertIndex + offset)}${row - 1},">0")`
};

export default formulas;
