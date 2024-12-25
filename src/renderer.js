import './styles.css';
import ExcelJS from 'exceljs';

let mainFilePath = '';
let avrFilePath = '';

const selectMainFile = document.getElementById('selectMainFile');
const selectAvrFile = document.getElementById('selectAvrFile');
const processFilesButton = document.getElementById('processFilesButton');
const outputDiv = document.getElementById('output');

const handleFileSelect = (setFilePath) => (event) => {
    const file = event.target.files[0];
    if (file) {
        setFilePath(file);
        console.log('Выбран файл:', file.name);
    }
};

const setMainFilePath = (file) => {
    mainFilePath = file;
};

const setAvrFilePath = (file) => {
    avrFilePath = file;
};

selectMainFile.addEventListener('change', handleFileSelect(setMainFilePath));
selectAvrFile.addEventListener('change', handleFileSelect(setAvrFilePath));

processFilesButton.addEventListener('click', async () => {
    if (mainFilePath && avrFilePath) {
        const processedData = await processExcelFiles(mainFilePath, avrFilePath);
        await saveFile(processedData);
    } else {
        alert('Пожалуйста, выберите оба файла.');
    }
});

async function processExcelFiles(mainFile, avrFile) {
    const mainWorkbook = new ExcelJS.Workbook();
    await mainWorkbook.xlsx.load(await mainFile.arrayBuffer());
    const mainSheet = mainWorkbook.worksheets[0];

    const avrWorkbook = new ExcelJS.Workbook();
    await avrWorkbook.xlsx.load(await avrFile.arrayBuffer());
    const avrSheet = avrWorkbook.worksheets[0];

    const avrFileName = avrFile.name.replace('.xlsx', '');
    const quantityColumnName = `Количество ${avrFileName}`;
    const costColumnName = `Стоимость ${avrFileName}`;

    const insertIndex = mainSheet.columnCount - 4;

    let quantityExists = false;
    let costExists = false;

    for (let col = 1; col <= mainSheet.columnCount; col++) {
        const headerCell = mainSheet.getCell(1, col);
        const header = headerCell.value;

        if (header === quantityColumnName) quantityExists = true;
        if (header === costColumnName) costExists = true;
    }

    if (!costExists) {
        mainSheet.spliceColumns(insertIndex, 0, [costColumnName]);
    }

    if (!quantityExists) {
        mainSheet.spliceColumns(insertIndex + (costExists ? 1 : 0), 0, [quantityColumnName]);
    }

    const defaultRowHeight = 14;
    const headerRowHeight = defaultRowHeight * 2;

    const textFormat = '@';
    const format = `_-* #,##0.00_-;_-* "-" #,##0.00_-;_-* "-"??_-;_-@_-`;
    const borderStyle = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
    };
    const headerFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE3F2FD' }
    };

    function calculateColumnWidth(sheet, col) {
        let maxWidth = 10;
        for (let row = 1; row <= sheet.rowCount; row++) {
            const cellValue = sheet.getCell(row, col).value;
            const cellWidth = cellValue ? String(cellValue).length : 0;
            maxWidth = Math.max(maxWidth, cellWidth);
        }
        return maxWidth + 1;
    }

    for (let row = 1; row <= mainSheet.rowCount; row++) {
        mainSheet.getRow(row).height = row === 1 ? headerRowHeight : defaultRowHeight;

        for (let col = 1; col <= mainSheet.columnCount; col++) {
            mainSheet.getColumn(col).width = col === 2 ? 40 : calculateColumnWidth(mainSheet, col);
            const cell = mainSheet.getCell(row, col);
            cell.style = {};

            if (cell.row === 1) {
                cell.fill = null;
                cell.fill = headerFill;
            }

            cell.numFmt = cell.col === 1 ? textFormat : format;
            cell.border = borderStyle;
        }
    }

    const mainData = mainSheet.getSheetValues().slice(2);
    const avrData = avrSheet.getSheetValues().slice(2);
    const avrMap = new Map(avrData.map(row => [row[2], row]));

    mainData.forEach((row, index) => {
        const avrRow = avrMap.get(row[2]);
        const cell = mainSheet.getCell(index + 2, insertIndex);
            cell.value = avrRow ? avrRow[4] : 0;
    });


    for (let row = 2; row <= mainData.length + 1; row++) {
        const lastColumnIndex = mainSheet.columnCount;
        const penultimateColumnIndex = lastColumnIndex - 1;
        const penultimateColumnLetter = String.fromCharCode(64 + penultimateColumnIndex);

        const columnLetterInsertIndexMinus1 = String.fromCharCode(65 + insertIndex - 1);
        const columnLetterInsertIndexPlus1 = String.fromCharCode(65 + insertIndex + 1);
        const columnLetterInsertIndexPlus3 = String.fromCharCode(65 + insertIndex + 3);

        const prevValue = mainSheet.getCell(`${columnLetterInsertIndexMinus1}${row}`).value;
        const curValue = mainSheet.getCell(`${columnLetterInsertIndexPlus1}${row}`).value;
        const totalValue = prevValue + curValue;

        const totalCostFormula = `D${row} * E${row}`;
        const avrCostFormula = `${columnLetterInsertIndexMinus1}${row} * E${row}`;
        const completedCostFormula = `${columnLetterInsertIndexPlus1}${row} * E${row}`;
        const quantityRemainingFormula = `D${row} - ${columnLetterInsertIndexPlus1}${row}`;
        const remainingCostFormula = `${columnLetterInsertIndexPlus3}${row} * E${row}`;
        const excessFormula = `IF(${penultimateColumnLetter}${row}<0, ABS(${penultimateColumnLetter}${row}), 0)`;

        mainSheet.getCell(`F${row}`).value = { formula: totalCostFormula };
        mainSheet.getCell(row, insertIndex + 1).value = { formula: avrCostFormula };
        mainSheet.getCell(row, insertIndex + 2).value = totalValue;
        mainSheet.getCell(row, insertIndex + 3).value = { formula: completedCostFormula };
        mainSheet.getCell(row, insertIndex + 4).value = { formula: quantityRemainingFormula };
        mainSheet.getCell(row, insertIndex + 5).value = { formula: remainingCostFormula };
        mainSheet.getCell(row, insertIndex + 6).value = { formula: excessFormula };
    }
    function getColumnLetter(columnIndex) {
        let letter = '';
        while (columnIndex > 0) {
            const modulo = (columnIndex - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            columnIndex = Math.floor((columnIndex - modulo) / 26);
        }
        console.log("letter: ", letter);
        return letter;
    }

    const lastColumnIndex = mainSheet.columnCount;

    if (lastColumnIndex > 0) {
        const penultimateColumnIndex = lastColumnIndex - 1;
        const penultimateColumnLetter = getColumnLetter(penultimateColumnIndex);

        console.log('Last Column Index:', lastColumnIndex);
        console.log('Penultimate Column Index:', penultimateColumnIndex);
        console.log('Penultimate Column Letter:', penultimateColumnLetter);

        const formula_Tomato = `=$${penultimateColumnLetter}2<0`;
        const formula_PastelGreen = `=AND(NOT(ISBLANK($${penultimateColumnLetter}2)),$${penultimateColumnLetter}2=0)`;

        try {
            const cellValue = mainSheet.getCell(`${penultimateColumnLetter}2`).value;
            console.log(`Value in ${penultimateColumnLetter}2:`, cellValue);

            mainSheet.addConditionalFormatting({
                ref: `A2:${getColumnLetter(lastColumnIndex)}${mainSheet.rowCount}`,
                rules: [
                    {
                        type: 'expression',
                        formulae: [formula_Tomato],
                        style: {
                            fill: {
                                type: 'pattern',
                                pattern: 'solid',
                                bgColor: { argb: 'FFFF6347' }
                            }
                        }
                    },
                    {
                        type: 'expression',
                        formulae: [formula_PastelGreen],
                        style: {
                            fill: {
                                type: 'pattern',
                                pattern: 'solid',
                                bgColor: { argb: 'C8FFC8' }
                            }
                        }
                    }

                ]
            });
        } catch (error) {
            console.error('Error adding conditional formatting:', error);
        }
    } else {
        console.error('No columns available in the sheet.');
    }
    return await mainWorkbook.xlsx.writeBuffer();
}

async function saveFile(data) {
    const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'data.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    URL.revokeObjectURL(url);
}
