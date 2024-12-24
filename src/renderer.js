import ExcelJS from 'exceljs';

let mainFilePath = '';
let avrFilePath = '';

document.getElementById('selectMainFile').addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (file) {
        mainFilePath = file;
        console.log('Выбран основной файл:', mainFilePath.name);
    }
});

document.getElementById('selectAvrFile').addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (file) {
        avrFilePath = file;
        console.log("avrFilePath:", avrFilePath);
        console.log('Выбран файл АВР:', avrFilePath.name);
    }
});

document.getElementById('processFilesButton').addEventListener('click', async () => {
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
        const header = mainSheet.getCell(1, col).value;
        if (header === quantityColumnName) quantityExists = true;
        if (header === costColumnName) costExists = true;
    }

    if (!costExists) {
        mainSheet.spliceColumns(insertIndex, 0, [costColumnName]);
    }

    if (!quantityExists) {
        mainSheet.spliceColumns(insertIndex + (costExists ? 1 : 0), 0, [quantityColumnName]);
    }

    const format = `_-* #,##0.00_-;_-* "-" #,##0.00_-;_-* "-"??_-;_-@_-`;
    for (let row = 1; row <= mainSheet.rowCount; row++) {
        for (let col = 1; col <= mainSheet.columnCount; col++) {
            mainSheet.getCell(row, col).numFmt = format;
        }
    }

    const borderStyle = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
    };

    const newColumns = [costColumnName, quantityColumnName];
    newColumns.forEach((column, idx) => {
        const colIndex = insertIndex + idx;
        for (let row = 1; row <= mainSheet.rowCount; row++) {
            const cell = mainSheet.getCell(row, colIndex);
            cell.border = {
                top: { style: borderStyle.top.style },
                left: { style: borderStyle.left.style },
                bottom: { style: borderStyle.bottom.style },
                right: { style: borderStyle.right.style },
            };
        }
    });

    const mainData = mainSheet.getSheetValues().slice(2);
    const avrData = avrSheet.getSheetValues().slice(2);
    const avrMap = new Map(avrData.map(row => [row[2], row]));

    mainData.forEach((row, index) => {
        const avrRow = avrMap.get(row[2]);
        if (avrRow) {
            mainSheet.getCell(index + 2, insertIndex).value = avrRow[4];
        }
    });

    for (let row = 2; row <= mainData.length + 1; row++) {
        const lastColumnIndex = mainSheet.columnCount;
        const penultimateColumnIndex = lastColumnIndex - 1;
        const costCompletedColumnIndex = lastColumnIndex - 3;
        const penultimateColumnLetter = String.fromCharCode(64 + penultimateColumnIndex);
        const costCompletedColumnLetter = String.fromCharCode(64 + costCompletedColumnIndex);
        console.log("costCompletedColumnLetter:", costCompletedColumnLetter);

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
