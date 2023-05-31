const fs = require('fs');
const path = require('path');

const { generateTableDoc } = require('./generateDocArrFuncPlus.js');

let fileJson = fs.readFileSync(path.join(__dirname, './file/json.json'), 'utf-8');
// console.log(6, fileJson);
let arr = JSON.parse(fileJson);

console.log(999, arr);

let outPutPath = `./result_${new Date().getTime()}.docx`

generateTableDoc(arr, outPutPath);

