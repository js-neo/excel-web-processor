import { excelConstants } from "../constants/index.js";
const { SCALE_FACTOR } = excelConstants;

export function getColumnLetter(columnIndex) {
    let letter = "";
    while (columnIndex > 0) {
        const modulo = (columnIndex - 1) % 26;
        letter = String.fromCharCode(65 + modulo) + letter;
        columnIndex = Math.floor((columnIndex - modulo) / 26);
    }
    return letter;
}

export function calculateCellWidth(value, font) {
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    context.font = `${font.size}px ${font.name}`;
    const metrics = context.measureText(value);
    return Math.round(metrics.width * SCALE_FACTOR);
}
