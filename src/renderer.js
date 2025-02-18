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

const {
    BASE_COLUMN_COUNT,
    QUANTITY_COLUMNS_COUNT,
    CHUNK_SIZE,
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
    console.log("event: ", event);
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
    const mainWorkbook = new ExcelJS.Workbook();
    await mainWorkbook.xlsx.load(await mainFile.arrayBuffer());
    const mainSheet = mainWorkbook.worksheets[0];

    const avrWorkbook = new ExcelJS.Workbook();
    await avrWorkbook.xlsx.load(await avrFile.arrayBuffer());
    const avrSheet = avrWorkbook.worksheets[0];

    console.log("avrSheet: ", avrSheet.getSheetValues());

    const avrFileName = avrFile.name.replace(".xlsx", "");
    const quantityColumnName = `${avrFileName} кол-во`;
    const costColumnName = `${avrFileName} сумма`;

    const insertIndex = mainSheet.columnCount - 4;

    const rowHeight = 14;
    const headerRowHeight = rowHeight * 2;

    let quantityExists = false;
    let costExists = false;

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

    const footerSecondCell = mainSheet.getCell(mainSheet.rowCount, 2);
    if (footerSecondCell.value === "Всего:") {
        mainSheet.spliceRows(mainSheet.rowCount, 1);
    }

    const mainKeys = new Set(
        mainSheet
            .getSheetValues()
            .slice(2)
            .map((row) => String(row[globals.processColNum] || "").trim())
    );

    const avrKeys = new Set(
        avrSheet
            .getSheetValues()
            .slice(2)
            .map((row) => String(row[globals.processColNum] || "").trim())
    );

    const allKeys = [...new Set([...mainKeys, ...avrKeys])]
        .filter((key) => key !== undefined)
        .sort(sortKeys);

    const newTable = [];
    const avrMap = new Map(
        avrSheet
            .getSheetValues()
            .slice(2)
            .map((row) => [row[globals.processColNum]?.trim(), row])
    );

    const mainValues = mainSheet.getSheetValues().slice(2);

    allKeys.forEach((key) => {
        console.log("Обработка строк таблицы");
        const mainRow = mainValues.find(
            (row) =>
                String(row[globals.processColNum]).trim() === String(key).trim()
        );

        if (mainRow) {
            mainRow[1] = mainRow[1] ? mainRow[1] : "";
            newTable.push(mainRow);
        } else {
            const avrRow = avrMap.get(key);
            const newRow = new Array(mainSheet.columnCount).fill(null);

            if (avrRow) {
                if (Number(globals.processColNum) === 1) {
                    newRow[0] = String(key);
                    newRow[1] = avrRow?.[2] || "";
                } else if (Number(globals.processColNum) === 2) {
                    newRow[0] = avrRow?.[1] || "";
                    newRow[1] = String(key);
                }

                newRow[2] = avrRow[3];
                newRow[3] = 0;
                newRow[4] = avrRow[5];
            } else {
                newRow[4] = 0;
            }

            newTable.push(newRow);
        }
    });
    console.log("Выход из цикла построения строк");

    while (mainSheet.rowCount > 1) {
        mainSheet.spliceRows(2, 1);
    }

    const remainingRows = mainSheet.getSheetValues().length - 2;

    if (remainingRows > 0) {
        console.error(`Ошибка: осталось ${remainingRows} строк в mainSheet.`);
    } else {
        newTable.forEach((row) => {
            console.log("Вставка новых строк");
            mainSheet.addRow(row);
        });
    }

    const totalColumns = mainSheet.columnCount;

    if (totalColumns >= 15) {
        const startColumn = 8;
        const endColumn = totalColumns - 7;
        console.log("Вставка формул в ячейки");
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

    for (let i = 1; i <= totalRows; i += CHUNK_SIZE) {
        await new Promise((resolve) => setTimeout(resolve, 0));
        console.log("Заполнение ячеек таблицы");

        const endRow = Math.min(i + CHUNK_SIZE - 1, totalRows);
        for (let row = i; row <= endRow; row++) {
            let maxHeight = rowHeight;

            for (let col = 1; col <= mainSheet.columnCount; col++) {
                const cell = mainSheet.getCell(row, col);
                const cellValue = cell.value;

                if ((col === 1 || col === 2) && row > 0) {
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
                            calculateCellWidth(
                                String(cellValue.toFixed(2)),
                                font
                            ) + EXTRA_WIDTH_FOR_NUMERIC;
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
                        String(cellValue).length / maxCellWidth
                    );
                    const cellHeight = numberOfLines * rowHeight;
                    maxHeight = Math.max(maxHeight, cellHeight);
                }
            }

            mainSheet.getRow(row).height =
                row === 1 ? headerRowHeight : maxHeight;
        }
    }

    for (let col = 1; col <= mainSheet.columnCount; col++) {
        mainSheet.getColumn(col).width = maxColumnWidths[col - 1];
    }

    const mainData = mainSheet.getSheetValues().slice(2);

    mainData.forEach((row, index) => {
        console.log("Заполнение колонок АВР");
        const avrRow = avrMap.get(row[globals.processColNum]);
        const cell = mainSheet.getCell(index + 2, insertIndex);
        cell.value = avrRow ? avrRow[4] : 0;
    });

    const lastRow = mainSheet.rowCount;
    const lastRowCellValue = mainSheet.getCell(lastRow, 2).value;

    if (lastRowCellValue !== "Всего:") {
        mainSheet.addRow(["", "Всего:"]);
        const newLastRowIndex = lastRow + 1;
        console.log("Построение строки Всего");
        for (let col = 1; col <= mainSheet.columnCount; col++) {
            const cell = mainSheet.getCell(newLastRowIndex, col);
            cell.style =
                col === 1 ? sheetStyle.footerTextStyle : sheetStyle.footerStyle;
        }

        // const insertIndex = mainSheet.columnCount - 4;
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

                console.log("headerValue: ", headerValue);

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

    // const BASE_COLUMN_COUNT = 11;
    // const QUANTITY_COLUMNS_COUNT = 2;

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
            console.log("Вставка формул в последние 5 колонок");
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
            console.log("trackingColumnIndex: ", trackingColumnIndex);
            console.log(
                "negativeValueCheck(trackingColumnIndex): ",
                negativeValueCheck(trackingColumnIndex)
            );
            console.log(
                "zeroValueCheck(trackingColumnIndex): ",
                zeroValueCheck(trackingColumnIndex)
            );

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
    return await mainWorkbook.xlsx.writeBuffer();
}
