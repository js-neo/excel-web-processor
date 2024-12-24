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

    console.log("Заголовки основного файла:");
    for (let col = 1; col <= mainSheet.columnCount; col++) {
        console.log(`Column ${col}: ${mainSheet.getCell(1, col).value}`);
    }

    const insertIndex = mainSheet.columnCount - 4;

    let quantityExists = false;
    let costExists = false;

    // Проверяем наличие колонок
    for (let col = 1; col <= mainSheet.columnCount; col++) {
        const header = mainSheet.getCell(1, col).value;
        if (header === quantityColumnName) quantityExists = true;
        if (header === costColumnName) costExists = true;
    }


    if (!costExists) {
        console.log(`Добавление колонки: ${costColumnName}`);
        mainSheet.spliceColumns(insertIndex, 0, [costColumnName]);
    }

    if (!quantityExists) {
        console.log(`Добавление колонки: ${quantityColumnName}`);
        mainSheet.spliceColumns(insertIndex + (costExists ? 1 : 0), 0, [quantityColumnName]);
    }

    console.log("Заголовки основного файла после добавления колонок:");
    for (let col = 1; col <= mainSheet.columnCount + (quantityExists ? 1 : 0) + (costExists ? 1 : 0); col++) {
        console.log(`Column ${col}: ${mainSheet.getCell(1, col).value}`);
    }

    // Обработка данных
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
        console.log("penultimateColumnLetter:", penultimateColumnLetter);
        console.log("costCompletedColumnLetter:", costCompletedColumnLetter);

        const totalCostFormula = `D${row} * E${row}`;
        const completedCostFormula = `G${row} * E${row}`;
        const remainingCostFormula = `D${row} - ${insertIndex}${row}`;


        const excessFormula = `ЕСЛИ(${penultimateColumnLetter}${row}<0; ABS(${penultimateColumnLetter}${row}); 0)`;
        console.log(`Excess formula: ${excessFormula}`);


        mainSheet.getCell(`F${row}`).value = { formula: totalCostFormula };
        mainSheet.getCell(`H${row}`).value = { formula: completedCostFormula };
        mainSheet.getCell(row, insertIndex + 2).value = { formula: remainingCostFormula };
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
