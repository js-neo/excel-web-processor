import "./styles.css";
import ExcelJS from "exceljs";

let mainFilePath = "";
let avrFilePath = "";
let processColNum = 1;

const mainFileInput = document.getElementById("selectMainFile");
const avrFileInput = document.getElementById("selectAvrFile");
const processFilesButton = document.getElementById("processFilesButton");
const outputDiv = document.getElementById("output");
const mainFileName = document.getElementById("mainFileName");
const avrFileName = document.getElementById("avrFileName");
const processColumnNumber = document.getElementById("processColumnNumber");

const handleFileSelect = (setFilePath) => (event) => {
  const file = event.target.files[0];
  if (file) {
    setFilePath(file);
  }
};

const setMainFilePath = (file) => {
  mainFilePath = file;
};

const setAvrFilePath = (file) => {
  avrFilePath = file;
};

const handleProcessColumnNumber = ({ target }) => {
  if (target.value > 0 && target.value <= 5) {
    processColNum = target.value;
  } else {
    processColNum = 1;
    target.value = processColNum;
  }
};

outputDiv.innerHTML =
  '<div class="info-message">Пожалуйста, загрузите основной файл сводной таблицы и файл АВР, ' +
  'затем нажмите кнопку "Обработать файлы".</div>';

mainFileInput.addEventListener("change", function (event) {
  mainFileName.textContent = this.files.length > 0 ? this.files[0].name : "";
  handleFileSelect(setMainFilePath)(event);
});

avrFileInput.addEventListener("change", function (event) {
  avrFileName.textContent = this.files.length > 0 ? this.files[0].name : "";
  handleFileSelect(setAvrFilePath)(event);
});

processFilesButton.addEventListener("click", async () => {
  if (mainFilePath && avrFilePath) {
    outputDiv.innerHTML =
      '<div class="processing-message">Идет обработка файлов.</div>';
    document.getElementById("timer").style.display = "block";

    const startTime = Date.now();
    let dotCount = 1;
    const maxDots = 5;

    const timerInterval = setInterval(() => {
      document.getElementById("timerValue").innerText = (
        (Date.now() - startTime) /
        1000
      ).toFixed(2);
    }, 100);

    const messageInterval = setInterval(() => {
      const dots = ".".repeat(dotCount);
      outputDiv.querySelector(".processing-message").textContent =
        `Идет обработка файлов${dots}`;
      dotCount = (dotCount % maxDots) + 1;
    }, 500);

    try {
      const processedData = await processExcelFiles(mainFilePath, avrFilePath);
      await saveFile(processedData);
      const endTime = Date.now();
      clearInterval(timerInterval);
      clearInterval(messageInterval);
      const duration = ((endTime - startTime) / 1000).toFixed(2);
      outputDiv.innerHTML = `<div class="success-message">Файлы успешно обработаны! 
Время выполнения: ${duration} секунд.</div>`;
      document.getElementById("timer").style.display = "none";
    } catch (error) {
      clearInterval(timerInterval);
      clearInterval(messageInterval);
      outputDiv.innerHTML = `<div class="error-message">Ошибка обработки файлов: ${error.message}. 
Пожалуйста, попробуйте снова.</div>`;
    } finally {
      document.getElementById("timer").style.display = "none";
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

  const avrFileName = avrFile.name.replace(".xlsx", "");
  const quantityColumnName = `${avrFileName} кол-во`;
  const costColumnName = `${avrFileName} сумма`;

  const insertIndex = mainSheet.columnCount - 4;
  const BASE_COLUMN_COUNT = 11;
  const QUANTITY_COLUMNS_COUNT = 2;
  const CHUNK_SIZE = 10;
  const STANDARD_DPI = 96;
  const BASE_SCALE_FACTOR = 400;

  let quantityExists = false;
  let costExists = false;

  const textFormat = "@";
  const format = `_-* #,##0.00_-;_-* "-" #,##0.00_-;_-* "-"??_-;_-@_-`;
  const borderStyle = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
  const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF87CEEB" },
  };

  const footerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFC107" },
  };

  const rowHeight = 14;
  const headerRowHeight = rowHeight * 2;

  const headerStyle = {
    font: {
      name: "Times New Roman",
      size: 10,
      bold: true,
    },
    alignment: {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    },
    fill: headerFill,
  };

  const contentStyle = {
    font: {
      name: "Arial",
      size: 9,
      bold: false,
    },
    alignment: {
      wrapText: true,
    },
  };

  const footerStyle = {
    font: {
      name: "Times New Roman",
      size: 10,
      bold: true,
    },
    alignment: {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    },
    fill: footerFill,
  };

  function getColumnLetter(columnIndex) {
    let letter = "";
    while (columnIndex > 0) {
      const modulo = (columnIndex - 1) % 26;
      letter = String.fromCharCode(65 + modulo) + letter;
      columnIndex = Math.floor((columnIndex - modulo) / 26);
    }
    return letter;
  }

  function calculateCellWidth(value, font) {
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    context.font = `${font.size}px ${font.name}`;
    const metrics = context.measureText(value);
    const dpi = window.devicePixelRatio * STANDARD_DPI;
    const scaleFactor = dpi / BASE_SCALE_FACTOR;
    return Math.round(metrics.width * scaleFactor);
  }

  for (let col = 1; col <= mainSheet.columnCount; col++) {
    const headerCell = mainSheet.getCell(1, col);
    const header = headerCell.value;

    if (header === quantityColumnName) quantityExists = true;
    if (header === costColumnName) costExists = true;
  }

  if (!costExists && !quantityExists) {
    mainSheet.spliceColumns(
      insertIndex,
      0,
      [quantityColumnName],
      [costColumnName],
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
      .map((row) => String(row[processColNum] || "").trim()),
  );

  const avrKeys = new Set(
    avrSheet
      .getSheetValues()
      .slice(2)
      .map((row) => String(row[processColNum] || "").trim()),
  );

  const allKeys = [...new Set([...mainKeys, ...avrKeys])]
    .filter((key) => key !== undefined)
    .sort((a, b) => {
      const splitRegex = /(\d+([.,]\d+)?)|([^0-9]+)/g;

      const getParts = (str) =>
        str.match(splitRegex)?.filter((x) => x !== undefined && x !== "") || [];

      const aParts = getParts(a);
      const bParts = getParts(b);

      for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
        const aPart = aParts[i] || "";
        const bPart = bParts[i] || "";

        const aIsNum = /^\d+([.,]\d+)?$/.test(aPart);
        const bIsNum = /^\d+([.,]\d+)?$/.test(bPart);

        if (aIsNum && bIsNum) {
          const aNum = parseFloat(aPart.replace(",", "."));
          const bNum = parseFloat(bPart.replace(",", "."));
          if (aNum !== bNum) return aNum - bNum;
        } else if (aIsNum !== bIsNum) {
          return aIsNum ? -1 : 1;
        } else {
          const diff = aPart.localeCompare(bPart, undefined, {
            sensitivity: "base",
          });
          if (diff !== 0) return diff;
        }
      }
      return 0;
    });

  const newTable = [];
  const avrMap = new Map(
    avrSheet
      .getSheetValues()
      .slice(2)
      .map((row) => [row[processColNum]?.trim(), row]),
  );

  const mainValues = mainSheet.getSheetValues().slice(2);

  allKeys.forEach((key) => {
    const mainRow = mainValues.find(
      (row) => String(row[processColNum]).trim() === String(key).trim(),
    );

    if (mainRow) {
      newTable.push(mainRow);
    } else {
      const avrRow = avrMap.get(key);
      const newRow = new Array(mainSheet.columnCount).fill(null);

      if (avrRow) {
        if (Number(processColNum) === 1) {
          newRow[0] = String(key);
          newRow[1] = avrRow?.[2] || "";
        } else if (Number(processColNum) === 2) {
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

  while (mainSheet.rowCount > 1) {
    mainSheet.spliceRows(2, 1);
  }

  const remainingRows = mainSheet.getSheetValues().length - 2;

  if (remainingRows > 0) {
    console.error(`Ошибка: осталось ${remainingRows} строк в mainSheet.`);
  } else {
    newTable.forEach((row) => {
      mainSheet.addRow(row);
    });
  }

  const totalColumns = mainSheet.columnCount;

  if (totalColumns >= 15) {
    const startColumn = 8;
    const endColumn = totalColumns - 7;

    for (let row = 2; row <= mainSheet.rowCount; row++) {
      for (let col = startColumn; col <= endColumn; col += 2) {
        const previousCellAddress = getColumnLetter(col - 1) + row;
        const formula = `=E${row}*${previousCellAddress}`;

        mainSheet.getCell(row, col).value = { formula: formula };
      }
    }
  }

  const maxColumnWidths = new Array(mainSheet.columnCount).fill(0);

  const totalRows = mainSheet.rowCount;

  for (let i = 1; i <= totalRows; i += CHUNK_SIZE) {
    await new Promise((resolve) => setTimeout(resolve, 0));

    const endRow = Math.min(i + CHUNK_SIZE - 1, totalRows);
    for (let row = i; row <= endRow; row++) {
      let maxHeight = rowHeight;

      for (let col = 1; col <= mainSheet.columnCount; col++) {
        const cell = mainSheet.getCell(row, col);
        const cellValue = cell.value;

        if (col === 1 && row > 1) {
          cell.value = String(cell.value).trim();
        }

        cell.style = {};

        let cellWidth = 0;
        const font = row === 1 ? headerStyle.font : contentStyle.font;

        if (row === 1) {
          cell.style = headerStyle;
        } else {
          cell.style = contentStyle;
        }
        cell.numFmt = col === 1 ? textFormat : format;
        cell.border = borderStyle;
        if (cellValue) {
          cellWidth = calculateCellWidth(String(cellValue), font);
        }
        const maxCellWidth = col === 2 ? 50 : 15;
        const effectiveCellWidth = Math.min(cellWidth, maxCellWidth);
        maxColumnWidths[col - 1] = Math.max(
          maxColumnWidths[col - 1],
          effectiveCellWidth,
        );

        if (cellValue) {
          const numberOfLines = Math.ceil(
            String(cellValue).length / maxCellWidth,
          );
          const cellHeight = numberOfLines * rowHeight;
          maxHeight = Math.max(maxHeight, cellHeight);
        }
      }

      mainSheet.getRow(row).height = row === 1 ? headerRowHeight : maxHeight;
    }
  }

  for (let col = 1; col <= mainSheet.columnCount; col++) {
    mainSheet.getColumn(col).width = maxColumnWidths[col - 1];
  }

  const mainData = mainSheet.getSheetValues().slice(2);

  mainData.forEach((row, index) => {
    const avrRow = avrMap.get(row[processColNum]);
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
      cell.style = footerStyle;
      cell.border = borderStyle;
      cell.numFmt = col === 1 ? textFormat : format;
    }

    const avrColumnOffset = 1;
    const completedColumnOffset = 3;
    const remainingColumnOffset = 5;
    const excessColumnOffset = 6;

    const costColumnIndex = (i) => insertIndex + i;
    const sumFormula = (num) =>
      `SUM(${getColumnLetter(costColumnIndex(num))}2:${getColumnLetter(costColumnIndex(num))}${newLastRowIndex - 1})`;
    mainSheet.getCell(`F${newLastRowIndex}`).value = {
      formula: `SUM(F2:F${newLastRowIndex - 1})`,
    };
    mainSheet.getCell(newLastRowIndex, costColumnIndex(avrColumnOffset)).value =
      { formula: sumFormula(avrColumnOffset) };
    mainSheet.getCell(
      newLastRowIndex,
      costColumnIndex(completedColumnOffset),
    ).value = { formula: sumFormula(completedColumnOffset) };
    mainSheet.getCell(
      newLastRowIndex,
      costColumnIndex(remainingColumnOffset),
    ).value = { formula: sumFormula(remainingColumnOffset) };
    mainSheet.getCell(
      newLastRowIndex,
      costColumnIndex(excessColumnOffset),
    ).value = { formula: sumFormula(excessColumnOffset) };
  }

  const calculateTotalQuantity = (rowIndex) => {
    const quantityColumnCount =
      (mainSheet.columnCount - BASE_COLUMN_COUNT) / QUANTITY_COLUMNS_COUNT;
    const firstQuantityColumnIndex = 7;
    return Array.from(
      { length: quantityColumnCount },
      (_, i) => firstQuantityColumnIndex + i * QUANTITY_COLUMNS_COUNT,
    )
      .map((colIndex) => `${getColumnLetter(colIndex)}${rowIndex}`)
      .join(",");
  };

  const lastColumnIndex = mainSheet.columnCount;
  if (lastColumnIndex > 0) {
    const penultimateColumnIndex = lastColumnIndex - 1;
    const trackingColumnIndex = lastColumnIndex - 2;
    const penultimateColumnLetter = getColumnLetter(penultimateColumnIndex);
    const trackingColumnLetter = getColumnLetter(trackingColumnIndex);

    for (let row = 2; row <= mainData.length + 1; row++) {
      const columnLetterInsertIndexMinus1 = getColumnLetter(insertIndex);
      const columnLetterInsertIndexPlus1 = getColumnLetter(insertIndex + 2);
      const columnLetterInsertIndexPlus3 = getColumnLetter(insertIndex + 4);

      const totalCostFormula = `D${row} * E${row}`;
      const avrCostFormula = `${columnLetterInsertIndexMinus1}${row} * E${row}`;
      const totalQuantityFormula = `SUM(${calculateTotalQuantity(row)})`;
      const completedCostFormula = `${columnLetterInsertIndexPlus1}${row} * E${row}`;
      const quantityRemainingFormula = `D${row} - ${columnLetterInsertIndexPlus1}${row}`;
      const remainingCostFormula = `${columnLetterInsertIndexPlus3}${row} * E${row}`;
      const excessFormula = `IF(${penultimateColumnLetter}${row}<0, ABS(${penultimateColumnLetter}${row}), 0)`;

      mainSheet.getCell(`F${row}`).value = { formula: totalCostFormula };
      mainSheet.getCell(row, insertIndex + 1).value = {
        formula: avrCostFormula,
      };
      mainSheet.getCell(row, insertIndex + 2).value = {
        formula: totalQuantityFormula,
      };
      mainSheet.getCell(row, insertIndex + 3).value = {
        formula: completedCostFormula,
      };
      mainSheet.getCell(row, insertIndex + 4).value = {
        formula: quantityRemainingFormula,
      };
      mainSheet.getCell(row, insertIndex + 5).value = {
        formula: remainingCostFormula,
      };
      mainSheet.getCell(row, insertIndex + 6).value = {
        formula: excessFormula,
      };
    }

    const formula_Tomato = `=$${trackingColumnLetter}2<0`;
    const formula_PastelGreen = `=AND(NOT(ISBLANK($${trackingColumnLetter}2)),$${trackingColumnLetter}2=0)`;

    try {
      const rangeRef = `A2:${getColumnLetter(lastColumnIndex)}${mainSheet.rowCount}`;
      mainSheet.removeConditionalFormatting(rangeRef);

      mainSheet.addConditionalFormatting({
        ref: `A2:${getColumnLetter(lastColumnIndex)}${mainSheet.rowCount}`,
        rules: [
          {
            type: "expression",
            formulae: [formula_Tomato],
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                bgColor: { argb: "FFFF6347" },
              },
            },
          },
          {
            type: "expression",
            formulae: [formula_PastelGreen],
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                bgColor: { argb: "C8FFC8" },
              },
            },
          },
        ],
      });
    } catch (error) {
      console.error("Ошибка при добавлении условного форматирования:", error);
    }
  }
  return await mainWorkbook.xlsx.writeBuffer();
}

async function saveFile(data) {
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "data.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  URL.revokeObjectURL(url);
}
