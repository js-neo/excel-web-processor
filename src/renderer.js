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
    console.log("avrFileName", avrFileName);
    const quantityColumnName = `Количество ${avrFileName}`;
    const costColumnName = `Стоимость ${avrFileName}`;

    const totalColumns = mainSheet.columnCount;
    const insertIndex = totalColumns - 4;
    console.log(`totalColumns: ${totalColumns}`);
    console.log(`insertIndex: ${insertIndex}`);

    let quantityExists = false;
    let costExists = false;

    for (let col = 1; col <= totalColumns; col++) {
        const header = mainSheet.getCell(1, col).value;
        console.log("header: ", header);
        if (header === quantityColumnName) quantityExists = true;
        if (header === costColumnName) costExists = true;
    }

    if (!quantityExists && !costExists) {
        mainSheet.spliceColumns(insertIndex, 0, [quantityColumnName] , [costColumnName]);
    }

    const mainData = mainSheet.getSheetValues().slice(2);
    const avrData = avrSheet.getSheetValues().slice(2);
    const avrMap = new Map(avrData.map(row => [row[1], row]));

    mainData.forEach((row, index) => {
        const avrRow = avrMap.get(row[1]);

        if (avrRow) {
            const quantityAvr = avrRow[2];
            const unitPrice = row[4];
            const doneColumnIndex = totalColumns - 5;
            const quantityDone = row[doneColumnIndex];

            if (!quantityExists && !costExists) {
                mainSheet.getCell(index + 2, insertIndex).value = quantityAvr;
                mainSheet.getCell(index + 2, insertIndex + 1).value = quantityAvr * unitPrice;
            }

            mainSheet.getCell(index + 2, doneColumnIndex).value = quantityDone + quantityAvr;
            mainSheet.getCell(index + 2, doneColumnIndex + 1).value = (quantityDone + quantityAvr) * unitPrice;

            const quantityRemaining = row[3] - (quantityDone + quantityAvr);
            mainSheet.getCell(index + 2, 9).value = quantityRemaining;
            mainSheet.getCell(index + 2, 10).value = quantityRemaining * unitPrice;

            const costRemaining = quantityRemaining * unitPrice;
            if (costRemaining < 0) {
                mainSheet.getCell(index + 2, 11).value = Math.abs(costRemaining);
            }
        }
    });

    for (let row = 2; row <= mainData.length + 1; row++) {
        mainSheet.getCell(`F${row}`).value = { formula: `D${row} * E${row}` };
        mainSheet.getCell(`H${row}`).value = { formula: `G${row} * E${row}` };
        mainSheet.getCell(`I${row}`).value = { formula: `D${row} - ${quantityColumnName}${row}` };
        mainSheet.getCell(`J${row}`).value = { formula: `I${row} * E${row}` };
        mainSheet.getCell(`K${row}`).value = { formula: `ЕСЛИ(J${row}<0; ABS(J${row}); 0)` };
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
