import "./styles/styles.css";
import ExcelJS from "exceljs";
import { excelConstants } from "./constants/index.js";
import { sortKeys } from "./utils/sortingUtils.js";
import { getColumnLetter, calculateCellWidth } from "./utils/excelUtils.js";
import { saveFile } from "./utils/fileSaver.js";
import { cellStyle } from "./styles/index.js";
import { formulas } from "./formulas/index.js";
import { globals } from "./dom/index.js";
import { fileHandlers } from "./handlers/index.js";
import { uiHandler } from "./ui/index.js";

class ProcessingError extends Error {
    constructor(messages, fileName) {
        const consoleMessage = `Ошибки в файле ${fileName}:\n${messages
            .map((error) => error.text)
            .join("\n")}`;
        super(consoleMessage);

        this.name = "ProcessingError";

        this.htmlMessage = `
            Ошибки в файле ${fileName}:<br>
            <ol>${messages.map((error) => `<li>${error.htmlElem}</li>`).join("")}</ol>
        `;
    }
}

const {
    BASE_COLUMN_COUNT,
    QUANTITY_COLUMNS_COUNT,
    NAME_COLUMN_WIDTH,
    BASE_COLUMN_WIDTH,
    EXTRA_WIDTH_FOR_NUMERIC
} = excelConstants;

const { sheetStyle, createFormattingOptions } = cellStyle;

const {
    totalCost,
    avrCost,
    totalQuantity,
    completedCost,
    quantityRemaining,
    remainingCost,
    excess,
    sum,
    sumIf,
    negativeValueCheck,
    zeroValueCheck
} = formulas;

const {
    mainFileInput,
    avrFileInput,
    processFilesButton,
    outputDiv,
    mainFileName,
    avrFileName,
    processColumnNumber,
    timerElement,
    timerValueElement
} = globals;

const {
    handleFileSelect,
    setMainFilePath,
    setAvrFilePath,
    handleProcessColumnNumber
} = fileHandlers;

const { showProcessingMessage, updateUI } = uiHandler;

updateUI(outputDiv, "info");

mainFileInput.addEventListener("change", function (event) {
    mainFileName.textContent = this.files.length > 0 ? this.files[0].name : "";
    handleFileSelect(setMainFilePath)(event);
});

avrFileInput.addEventListener("change", function (event) {
    avrFileName.textContent = this.files.length > 0 ? this.files[0].name : "";
    handleFileSelect(setAvrFilePath)(event);
});

processFilesButton.addEventListener("click", async () => {
    if (globals.mainFilePath && globals.avrFilePath) {
        const processingUI = showProcessingMessage(
            outputDiv,
            timerElement,
            timerValueElement
        );
        try {
            const processedData = await processExcelFiles(
                globals.mainFilePath,
                globals.avrFilePath
            );
            await saveFile(processedData);
            const duration = processingUI.getDuration();
            updateUI(outputDiv, "success", { duration });
        } catch (error) {
            updateUI(outputDiv, "error", { errorMessage: error.message });
        } finally {
            processingUI.stop();
        }
    } else {
        alert("Пожалуйста, выберите оба файла.");
    }
});

processColumnNumber.addEventListener("change", handleProcessColumnNumber);

async function processExcelFiles(mainFile, avrFile) {
    console.time("Initial table data");
    const mainWorkbook = new ExcelJS.Workbook();
    await mainWorkbook.xlsx.load(await mainFile.arrayBuffer());
    let mainSheet = mainWorkbook.worksheets[0];

    const avrWorkbook = new ExcelJS.Workbook();
    await avrWorkbook.xlsx.load(await avrFile.arrayBuffer());
    const avrSheet = avrWorkbook.worksheets[0];

    console.timeEnd("Initial table data");

    const avrFileName = avrFile.name.replace(".xlsx", "");
    const quantityColumnName = `${avrFileName} кол-во`;
    const costColumnName = `${avrFileName} сумма`;

    const insertIndex = mainSheet.columnCount - 4;

    const rowHeight = 28;
    const headerRowHeight = rowHeight * 2;

    let quantityExists = false;
    let costExists = false;

    console.time("Insert new column");

    for (let col = 1; col <= mainSheet.columnCount; col++) {
        const headerCell = mainSheet.getCell(1, col);
        const header = String(headerCell.value).trim();

        if (header === quantityColumnName) quantityExists = true;
        if (header === costColumnName) costExists = true;
    }

    if (!costExists && !quantityExists) {
        mainSheet.spliceColumns(
            insertIndex,
            0,
            [quantityColumnName],
            [costColumnName]
        );
    }

    console.timeEnd("Insert new column");

    const footerSecondCell = mainSheet.getCell(mainSheet.rowCount, 2);
    if (footerSecondCell.value === "Всего:") {
        mainSheet.spliceRows(mainSheet.rowCount, 1);
    }
    const mainMap = new Map();
    const avrMap = new Map();
    const getErrorMessage = ({
        processCol,
        excelRowIndex,
        typeErrorValue
    }) => ({
        text: `в ячейке ${getColumnLetter(processCol)}${excelRowIndex} обнаружено ${typeErrorValue} значение`,
        htmlElem: `в ячейке <strong>${getColumnLetter(processCol)}${excelRowIndex}</strong> 
        обнаружено ${typeErrorValue} значение`
    });
    const populateMapAndGetKeys = async (
        sheet,
        targetMap,
        processCol,
        fileName
    ) => {
        const rows = sheet.getSheetValues().slice(2);
        if (!Array.isArray(rows)) {
            throw new Error(`Данные листа не являются массивом!`);
        }
        const CHUNK_SIZE = Math.max(50, Math.floor(rows.length / 100));

        const keys = [];
        const errors = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const rawValue = row[processCol];
            const excelRowIndex = i + 2;
            const strValue = String(rawValue ?? "").trim();
            if (strValue === "") {
                errors.push(
                    getErrorMessage({
                        processCol,
                        excelRowIndex,
                        typeErrorValue: "пустое"
                    })
                );
                continue;
            }
            if (targetMap.has(strValue)) {
                errors.push(
                    getErrorMessage({
                        processCol,
                        excelRowIndex,
                        typeErrorValue: "дублирующее"
                    })
                );
                continue;
            }
            targetMap.set(strValue, row);
            keys.push(strValue);
            if (i % CHUNK_SIZE === 0)
                await new Promise((resolve) => setTimeout(resolve, 0));
        }
        if (errors.length > 0) {
            throw new ProcessingError(errors, fileName);
        }
        return keys;
    };
    console.time("Build Map rows and all keys");
    let allKeys = [];
    try {
        const mainKeys = await populateMapAndGetKeys(
            mainSheet,
            mainMap,
            globals.processColNum,
            "сводной таблицы"
        );
        const avrKeys = await populateMapAndGetKeys(
            avrSheet,
            avrMap,
            globals.processColNum,
            "АВР"
        );

        allKeys = [...new Set([...mainKeys, ...avrKeys])].sort(sortKeys);
    } catch (error) {
        if (error instanceof ProcessingError) {
            console.log(error.message);
            throw new Error(
                `Ошибка создания массивов ключей: <div style="margin-left: 20px;"> ${error.htmlMessage}</div>`
            );
        } else {
            throw new Error(
                `Ошибка создания массивов ключей: <div style="margin-left: 20px;"> ${error.message}</div>`
            );
        }
    }

    console.timeEnd("Build Map rows and all keys");

    console.time("Build rows");
    const newTable = await buildRows(allKeys, mainMap, avrMap, mainSheet);
    console.timeEnd("Build rows");

    console.time("Deleted rows");

    if (mainSheet.rowCount > 1) {
        const headers = mainSheet.getRow(1).values;
        mainSheet.spliceRows(2, mainSheet.rowCount - 1);

        if (mainSheet.rowCount > 1) {
            const newSheet = mainWorkbook.addWorksheet("Temp Sheet");
            newSheet.addRow(headers);
            mainWorkbook.removeWorksheet(mainSheet.id);
            newSheet.name = mainSheet.name;
            mainSheet = newSheet;
        }
    }

    console.timeEnd("Deleted rows");

    const remainingRows = mainSheet.rowCount - 1;

    if (remainingRows > 0) {
        console.error(`Ошибка: осталось ${remainingRows} строк в mainSheet.`);
    } else {
        console.log("Все строки успешно удалены.");

        console.time("Insert new rows");

        const filteredTable = newTable.filter((row) =>
            row.slice(1, 3).some((cell) => cell !== null && cell !== "")
        );
        mainSheet.addRows(filteredTable);
        console.timeEnd("Insert new rows");
    }

    const totalColumns = mainSheet.columnCount;

    if (totalColumns >= 15) {
        const startColumn = 8;
        const endColumn = totalColumns - 7;
        for (let row = 2; row <= mainSheet.rowCount; row++) {
            for (let col = startColumn; col <= endColumn; col += 2) {
                const prevColIndex = col - 1;
                mainSheet.getCell(row, col).value = {
                    formula: avrCost(row, prevColIndex)
                };
            }
        }
    }

    const maxColumnWidths = new Array(mainSheet.columnCount).fill(0);

    const totalRows = mainSheet.rowCount;

    const mainData = mainSheet.getSheetValues().slice(2);

    mainData.forEach((row, index) => {
        const avrRow = avrMap.get(row[globals.processColNum].trim());
        const cell = mainSheet.getCell(index + 2, insertIndex);
        cell.value = avrRow ? avrRow[4] : 0;
    });

    const lastRow = mainSheet.rowCount;
    const lastRowCellValue = mainSheet.getCell(lastRow, 2).value;

    if (lastRowCellValue !== "Всего:") {
        mainSheet.addRow(["", "Всего:"]);
        const newLastRowIndex = lastRow + 1;

        for (let col = 1; col <= mainSheet.columnCount; col++) {
            const cell = mainSheet.getCell(newLastRowIndex, col);
            cell.style =
                col === 1 ? sheetStyle.footerTextStyle : sheetStyle.footerStyle;
        }

        const remainingAmountOffset = 5;
        const excessAmountOffset = 6;

        const costColumnIndex = (i) => insertIndex + i;

        if (totalColumns >= 11) {
            const startColumn = 6;
            const endColumn = totalColumns - 3;

            for (let col = startColumn; col <= endColumn; col += 2) {
                const columnLetter = getColumnLetter(col);
                const arrHeaderValues = mainSheet
                    .getCell(1, col)
                    .value.split(" ");

                const headerValue =
                    arrHeaderValues.length > 1
                        ? arrHeaderValues
                              .slice(0, arrHeaderValues.length - 1)
                              .join(" ")
                        : arrHeaderValues[0];

                mainSheet.getCell(newLastRowIndex, col - 1).value =
                    `${headerValue}:`;
                const formula = `=SUM(${columnLetter}2:${columnLetter}${newLastRowIndex - 1})`;

                mainSheet.getCell(newLastRowIndex, col).value = {
                    formula: formula
                };
            }
        }

        mainSheet.getCell(
            newLastRowIndex,
            costColumnIndex(remainingAmountOffset)
        ).value = {
            formula: sumIf(newLastRowIndex, insertIndex, remainingAmountOffset)
        };
        mainSheet.getCell(
            newLastRowIndex,
            costColumnIndex(excessAmountOffset)
        ).value = {
            formula: sum(newLastRowIndex, insertIndex, excessAmountOffset)
        };
    }

    const calculateTotalQuantity = (rowIndex) => {
        const quantityColumnCount =
            (mainSheet.columnCount - BASE_COLUMN_COUNT) /
            QUANTITY_COLUMNS_COUNT;
        const firstQuantityColumnIndex = 7;
        return Array.from(
            {
                length: quantityColumnCount
            },
            (_, i) => firstQuantityColumnIndex + i * QUANTITY_COLUMNS_COUNT
        )
            .map((colIndex) => `${getColumnLetter(colIndex)}${rowIndex}`)
            .join(",");
    };

    const lastColumnIndex = mainSheet.columnCount;
    if (lastColumnIndex > 0) {
        const penultimateColumnIndex = lastColumnIndex - 1;
        const trackingColumnIndex = lastColumnIndex - 2;
        const penultimateColumnLetter = getColumnLetter(penultimateColumnIndex);

        for (let row = 2; row <= mainData.length + 1; row++) {
            mainSheet.getCell(`F${row}`).value = {
                formula: totalCost(row)
            };
            mainSheet.getCell(row, insertIndex + 1).value = {
                formula: avrCost(row, insertIndex)
            };
            mainSheet.getCell(row, insertIndex + 2).value = {
                formula: totalQuantity(row, calculateTotalQuantity)
            };
            mainSheet.getCell(row, insertIndex + 3).value = {
                formula: completedCost(row, insertIndex)
            };
            mainSheet.getCell(row, insertIndex + 4).value = {
                formula: quantityRemaining(row, insertIndex)
            };
            mainSheet.getCell(row, insertIndex + 5).value = {
                formula: remainingCost(row, insertIndex)
            };
            mainSheet.getCell(row, insertIndex + 6).value = {
                formula: excess(row, penultimateColumnLetter)
            };
        }

        try {
            const rowCount = mainSheet.rowCount;
            const lastColumnLetter = getColumnLetter(lastColumnIndex);
            const rangeRef = `A2:${lastColumnLetter}${rowCount}`;
            mainSheet.removeConditionalFormatting(rangeRef);

            const formattingOptions = createFormattingOptions(
                rangeRef,
                negativeValueCheck(trackingColumnIndex),
                zeroValueCheck(trackingColumnIndex)
            );

            mainSheet.addConditionalFormatting(formattingOptions);
        } catch (error) {
            console.error(
                "Ошибка при добавлении условного форматирования:",
                error
            );
        }
    }

    console.time("applyStyles");

    for (let row = 1; row <= totalRows; row++) {
        let maxHeight = rowHeight;

        for (let col = 1; col <= mainSheet.columnCount; col++) {
            const cell = mainSheet.getCell(row, col);
            const cellValue = cell.value;

            if ((col === 1 || col === 2) && row > 1) {
                cell.value = String(cell.value).trim();
            }

            cell.style = {};
            if (row === 1) {
                cell.style = sheetStyle.headerStyle;
            } else {
                cell.style =
                    col === 1
                        ? sheetStyle.contentTextStyle
                        : sheetStyle.contentStyle;
            }

            let cellWidth = 0;
            const font =
                row === 1
                    ? sheetStyle.headerStyle.font
                    : sheetStyle.contentStyle.font;

            if (cellValue) {
                if (typeof cellValue === "number") {
                    cellWidth =
                        calculateCellWidth(String(cellValue.toFixed(2)), font) +
                        EXTRA_WIDTH_FOR_NUMERIC;
                } else {
                    cellWidth = calculateCellWidth(String(cellValue), font);
                }
            }

            const maxCellWidth =
                col === 2 ? NAME_COLUMN_WIDTH : BASE_COLUMN_WIDTH;
            const effectiveCellWidth = Math.min(cellWidth, maxCellWidth);

            maxColumnWidths[col - 1] = Math.max(
                maxColumnWidths[col - 1],
                effectiveCellWidth
            );

            if (cellValue) {
                const numberOfLines = Math.ceil(
                    (String(cellValue).length - 10) / maxCellWidth
                );
                const cellHeight = numberOfLines * rowHeight;
                maxHeight = Math.max(maxHeight, cellHeight);
            }
        }

        mainSheet.getRow(row).height = row === 1 ? headerRowHeight : maxHeight;

        if (row % 30 === 0)
            await new Promise((resolve) => setTimeout(resolve, 0));
    }
    console.timeEnd("applyStyles");

    for (let col = 1; col <= mainSheet.columnCount; col++) {
        mainSheet.getColumn(col).width = maxColumnWidths[col - 1];
    }
    return await mainWorkbook.xlsx.writeBuffer();
}

async function buildRows(allKeys, mainMap, avrMap, mainSheet) {
    const CHUNK_SIZE = 100;
    const processColIdx = globals.processColNum;

    const processRow = (key) => {
        const mainRow = mainMap.get(key);

        if (mainRow) {
            mainRow[1] = mainRow[1] || "";
            return mainRow;
        }
        const avrRow = avrMap.get(key);
        const newRow = new Array(mainSheet.columnCount).fill(null);

        if (avrRow) {
            newRow[0] =
                Number(processColIdx) === 1 ? String(key) : avrRow[1] || "";
            newRow[1] =
                Number(processColIdx) === 2 ? String(key) : avrRow[2] || "";
            newRow[2] = avrRow[3];
            newRow[3] = 0;
            newRow[4] = avrRow[5];
        } else {
            newRow[4] = 0;
        }

        return newRow;
    };

    const processChunk = async (chunk, chunkIndex) => {
        const chunkResult = [];
        for (const key of chunk) {
            try {
                const row = processRow(key);
                if (row) {
                    chunkResult.push(row);
                }
            } catch (error) {
                console.error(`Error processing key ${key}: ${error}`);
            }
        }
        if (chunkIndex % 10 === 0) {
            await new Promise((resolve) => setTimeout(resolve, 0));
        }

        return chunkResult;
    };

    const chunks = [];

    for (let i = 0; i < allKeys.length; i += CHUNK_SIZE) {
        const chunk = allKeys.slice(i, i + CHUNK_SIZE);
        chunks.push(chunk);
    }

    const results = await Promise.all(
        chunks.map((chunk, index) => processChunk(chunk, index))
    );

    return results.flat();
}
