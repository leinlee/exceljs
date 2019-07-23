const Workbook = require('../lib/doc/workbook');

const wb = new Workbook();
const sheet = wb.addWorksheet('TestSheet');
sheet.addRows([
    [1111],
    [2222],
    [3333],
    [4444],
    [5555],
]);

console.log('expects 5, got', sheet.rowCount);
sheet.spliceRows(2, 1);
console.log('expects 4, got', sheet.rowCount);
sheet.spliceRows(2, 0, ['+6666'], ['+7777']);
console.log('expects 6, got', sheet.rowCount);
sheet.spliceRows(2, sheet.rowCount);
console.log('expects 1, got', sheet.rowCount);
