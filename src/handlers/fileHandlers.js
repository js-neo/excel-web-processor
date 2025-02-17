import { globals } from "../dom/index.js";

let { mainFilePath, avrFilePath, processColNum } = globals;

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
    processColNum = target.value;
};

export default {
    handleFileSelect,
    setMainFilePath,
    setAvrFilePath,
    handleProcessColumnNumber
};
