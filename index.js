const fileSystem = require('fs');
const excelLibrary = require('xlsx');
const inputJsonFile = 'input.json';


// functions
const generateNestedWorkbook = (dataObj, parentKey = '', defaultVal = '') => {
    let resultObj = {};
    for (let key in dataObj) {
        if (typeof dataObj[key] === 'object' && dataObj[key] !== null) {
            const nestedSheetName = `${parentKey}${key}`;
            resultObj[parentKey + key] = nestedSheetName;
            const nestedData = generateNestedWorkbook(dataObj[key], '', defaultVal);
            appendNestedSheet(nestedData, nestedSheetName);
        } else
            resultObj[`${parentKey}${key}`] = dataObj[key] !== null ? dataObj[key] : defaultVal;
    }
    return resultObj;
};

const appendNestedSheet = (nestedData, sheetName) => {
    sheetName = String(sheetName);
    let index = 1;
    let modifiedSheetName = sheetName;
    while (nestedWorkbook.SheetNames.indexOf(modifiedSheetName) >= 0)
        modifiedSheetName = `${sheetName}_${index++}`;
    const worksheet = excelLibrary.utils.json_to_sheet([nestedData]);
    const nestedSheet = excelLibrary.utils.aoa_to_sheet(excelLibrary.utils.sheet_to_json(worksheet, { header: 1 }));
    excelLibrary.utils.book_append_sheet(nestedWorkbook, nestedSheet, modifiedSheetName);
};


// try catch block
try {
    const jsonData = JSON.parse(fileSystem.readFileSync(inputJsonFile, 'utf8'));
    const mainWorkbook = excelLibrary.utils.book_new();
    var nestedWorkbook = excelLibrary.utils.book_new();
    const flattenedData = generateNestedWorkbook(jsonData, '', '');
    const mainSheet = excelLibrary.utils.json_to_sheet([flattenedData]);
    excelLibrary.utils.book_append_sheet(mainWorkbook, mainSheet, 'MainSheet');
    const nestedOutputPath = 'nested_output.xlsx';
    excelLibrary.writeFile(nestedWorkbook, nestedOutputPath, { bookType: 'xlsx', bookSST: false, type: 'file' });
    const mainOutputPath = 'main_output.xlsx';
    excelLibrary.writeFile(mainWorkbook, mainOutputPath, { bookType: 'xlsx', bookSST: false, type: 'file' });
} catch (error) {
    console.error(`Error: ${error.message}`);
}
