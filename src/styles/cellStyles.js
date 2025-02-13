const textFormat = "@";
const format = `_-* #,##0.00_-;_-* "-" #,##0.00_-;_-* "-"??_-;_-@_-`;
const borderStyle = {
    top: {
        style: "thin"
    },
    left: {
        style: "thin"
    },
    bottom: {
        style: "thin"
    },
    right: {
        style: "thin"
    }
};
const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
        argb: "FFD2EBFA"
    }
};

const footerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: {
        argb: "FF8CBAD"
    }
};

export const sheetStyle = {
    headerStyle: {
        font: {
            name: "Times New Roman",
            size: 11,
            bold: true
        },
        alignment: {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
        },
        fill: headerFill,
        border: borderStyle,
        numFmt: textFormat
    },
    contentTextStyle: {
        font: {
            name: "Arial",
            size: 9,
            bold: false
        },
        alignment: {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
        },
        border: borderStyle,
        numFmt: textFormat
    },
    contentStyle: {
        font: {
            name: "Arial",
            size: 9,
            bold: false
        },
        alignment: {
            wrapText: true
        },
        border: borderStyle,
        numFmt: format
    },
    footerTextStyle: {
        font: {
            name: "Times New Roman",
            size: 10,
            bold: true
        },
        alignment: {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
        },
        fill: footerFill,
        border: borderStyle,
        numFmt: textFormat
    },
    footerStyle: {
        font: {
            name: "Times New Roman",
            size: 10,
            bold: true
        },
        alignment: {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
        },
        fill: footerFill,
        border: borderStyle,
        numFmt: format
    }
};
