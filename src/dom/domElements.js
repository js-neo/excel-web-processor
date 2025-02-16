const mainFileInput = document.getElementById("selectMainFile");
const avrFileInput = document.getElementById("selectAvrFile");
const processFilesButton = document.getElementById("processFilesButton");
const outputDiv = document.getElementById("output");
const mainFileName = document.getElementById("mainFileName");
const avrFileName = document.getElementById("avrFileName");
const processColumnNumber = document.getElementById("processColumnNumber");

let mainFilePath = "";
let avrFilePath = "";
let processColNum = 1;

export default {
    mainFileInput,
    avrFileInput,
    processFilesButton,
    outputDiv,
    mainFileName,
    avrFileName,
    processColumnNumber,
    mainFilePath,
    avrFilePath,
    processColNum
};
