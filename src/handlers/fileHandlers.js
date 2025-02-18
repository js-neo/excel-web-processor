import { globals } from "../dom/index.js";

const handleFileSelect = (setFilePath) => (event) => {
    const file = event.target.files[0];
    if (file) {
        setFilePath(file);
    }
};

const setMainFilePath = (file) => {
    globals.mainFilePath = file;
};

const setAvrFilePath = (file) => {
    globals.avrFilePath = file;
};

const handleProcessColumnNumber = ({ target }) => {
    globals.processColNum = target.value;
};

export default {
    handleFileSelect,
    setMainFilePath,
    setAvrFilePath,
    handleProcessColumnNumber
};
