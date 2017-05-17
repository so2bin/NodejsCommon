'use strict';

const fs = require('fs');
const log4js = require('log4js');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');

const logger = log4js.getLogger();

// excel cell default value if empty
const DEFAULT_CELL = "";

class XlsxRowReader {
    /**
     * read xls/xlsx file synchromously, and the column range must less than 'Z',
     * the row num is not limited,
     * the empty row will be ignored
     * 
     * @param  {String}  filepath : file path
     * @param  {Boolean} bHeader  : whther the first line is header
     */
    constructor(filepath, bHeader = true) {
        try {
            fs.accessSync(filepath, fs.F_OK | fs.R_OK);
        } catch (err) {
            throw new Error(`file is not accessed: ${filepath}, ${err.message}`);
        }
        logger.info(`read file ${filepath} succeed!`);
        this.workbook = XLSX.readFile(filepath);
        let sheetNames = this.workbook.SheetNames;
        this.sheet1 = this.workbook.Sheets[sheetNames[0]];
        if (!this.sheet1["!ref"]) {
            throw new Error('empty excel file');
        }

        // data range
        let eDtRng = this.sheet1["!ref"].split(':');
        this.col_s_char = eDtRng[0].match(/[A-Z]+/)[0];
        this.col_e_char = eDtRng[1].match(/[A-Z]+/)[0];
        this.row_s_num = eDtRng[0].match(/[0-9]+/)[0];
        this.row_e_num = eDtRng[1].match(/[0-9]+/)[0];
        if (this.col_s_char.length > 1 || this.col_e_char.length > 1) {
            throw new Error(`data column range must agree: col < "Z", now column: ${e_char}`);
        }
        if (bHeader) {
            ++this.row_s_num;
        }
        logger.info(`data range: ${eDtRng}`);
    }

    /**
     * read the xls/xlsx file with: for(let row_data of xlsxrowreader) {....},
     * row_data: [5, 'heli', '22'], in the array, the first value: 5 is row number, 
     * uoyi
     * and the left is the data in sequence
     **/
    * [Symbol.iterator]() {
        for (let row = this.row_s_num; row <= this.row_e_num; ++row) {
            let rowArr = [row];
            let bEmptyRow = true;
            for (let col = this.col_s_char.charCodeAt(); col <= this.col_e_char.charCodeAt(); ++col) {
                let pos = String.fromCharCode(col) + row;
                let cell = this.sheet1[pos] ? this.sheet1[pos].v : DEFAULT_CELL;
                if (cell !== DEFAULT_CELL) {
                    bEmptyRow = false;
                }
                rowArr.push(String(cell));
            }
            if (bEmptyRow) {
                continue;
            }
            yield rowArr;
        }
    }
}

module.exports = XlsxRowReader;

if (!module.parent) {
    let xlsxRdr = new XlsxRowReader('./spiker.xls');
    for (let row of xlsxRdr) {
        console.log(row)
    }
}
