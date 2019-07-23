'use strict';

const HrStopwatch = require('./utils/hr-stopwatch');
const Excel = require('../excel');

const fileIn = process.argv[2];
const fileOut = process.argv[3];
const wb = new Excel.Workbook();

const stopwatch = new HrStopwatch();
stopwatch.start();

wb.xlsx
  .readFile(fileIn)
  .then(() => {
    const micros = stopwatch.microseconds;

    console.log('Loaded', fileIn);
    console.log('Time taken:', micros / 1000000);

    const sheet = wb.getWorksheet('site_map');
    // console.log(sheet.rowCount)
    sheet.spliceRows(2, 3, ['1100']);
    // console.log(sheet.rowCount)
    wb.xlsx.writeFile(fileOut);
  })
  .catch(error => {
    console.error('something went wrong', error.stack);
  });
