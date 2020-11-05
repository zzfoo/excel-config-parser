const exceljs = require('exceljs');
const async = require('async');
const config = require('./config');

function parseExcel(filePath, callback) {
    console.log('===================================');
    console.log('===================================');
    console.log('===================================');
    console.log('parsing excel: ', filePath);
    const data = {};
    const workbook = new exceljs.Workbook();
    workbook.xlsx.readFile(filePath)
        .then(function (onfullfilled, onrejected) {
            const worksheets = workbook.worksheets;
            try {
                worksheets.forEach(function (sheet) {
                    const name = sheet.name;
                    console.log('===================================');
                    console.log('parsing sheet: ', name);
                    const sheetData = parseSheet(sheet);
                    // console.log('sheetData: ', sheetData);
                    data[name] = sheetData;
                })

            } catch (e) {
                console.log(e)
                onrejected(e)
            }
            // console.log('data: ', data);
            callback(null, data);
        })
        .catch(function (e) {
            callback(e);
        })
}

function parseExcels(filePaths, callback) {
    async.mapSeries(filePaths, parseExcel, function (err, datas) {
        callback(err, datas);
    })
}

function parseSheet(worksheet) {
    const firstRow = worksheet.getRow(1);
    if (firstRow.getCell(1).value === 'key' && firstRow.getCell(2).value === 'value' && firstRow.getCell(3).value === 'type') {
        // console.log('type: map');
        return parseMap(worksheet);
    }
    // console.log('type: array');
    return parseArray(worksheet);
}

function parseMap(worksheet) {
    const data = {};
    let key;
    let value;
    let type;
    worksheet.eachRow(function (row, rowNumber) {
        if (rowNumber <= 1) {
            return;
        }
        key = row.getCell(1).text;
        value = row.getCell(2).text.trim();
        type = row.getCell(3).text;
        if (!key || !type) {
            throw new Error('worksheet: ' + worksheet.name + ' ' + rowNumber + ' key or type is empty!');
        }
        try {
            data[key] = parseValue(value, type);
        } catch (e) {
            console.error('worksheet: ' + worksheet.name + ' row: ' + rowNumber + ' parse failed! ');
            console.error('key: ' + key + ' , value: ' + value + ' , type: ' + type);
            throw e;
        }
    })
    return data;
}

function parseArray(worksheet) {
    const data = [];

    const keys = {};
    const options = config.getOptions()
    const startRow = options.startRow || 1
    const startCol = options.startCol || 1
    const invalidRowMark = options.invalidRowMark
    worksheet.getRow(startRow).eachCell(function (cell, cellNumber) {
        if (cellNumber < startCol) return;
        keys[cellNumber] = cell.text.trim();
    })

    // console.log('PARSE_ARRAY: ', keys)
    const hasExplicitType = checkArrayHasExplicitType(worksheet, startRow, startCol)
    const dataStartRow = startRow + (hasExplicitType ? 2 : 1)

    const types = {};
    if (!hasExplicitType) {
        worksheet.getRow(startRow).eachCell(function (cell, cellNumber) {
            if (cellNumber < startCol || !keys[cellNumber]) return;

            types[cellNumber] = getColType(worksheet, cellNumber, dataStartRow)
        })
    } else {
        worksheet.getRow(startRow + 1).eachCell(function (cell, cellNumber) {
            if (cellNumber < startCol || !keys[cellNumber]) return;

            types[cellNumber] = cell.text.trim()
        })
    }
    // console.log(types)
    // return data

    worksheet.eachRow(function (row, rowNumber) {
        if (rowNumber < dataStartRow) {
            return;
        }

        if (invalidRowMark && row.getCell(1).text === invalidRowMark) {
            return
        }
        const item = {};
        row.eachCell(function (cell, cellNumber) {
            const key = keys[cellNumber];
            const type = types[cellNumber];
            if (!key || !type) return;

            const value = cell.text.trim();
            try {
                item[key] = parseValue(value, type);
            } catch (e) {
                console.log(e);
                console.error('worksheet: ' + worksheet.name + ' row: ' + rowNumber + ' parse failed! ');
                console.error('key: ' + key + ' , value: ' + value + ' , type: ' + type);
                throw e;
            }
        })
        data.push(item);
    })
    return data;
}

function parseValue(value, type) {
    if (value === '') {
        return undefined
    }
    if (typeof value === 'number') {
        value = value.toString()
    }
    switch (type) {
        case 'int':
            return parseInt(value);
        case 'num':
        case 'float':
        case 'number':
            return Number(value);
        case 'time':
            const r = /(-?)(?:(\d+)d)?(?:(\d+)h)?(?:(\d+)m)?(?:(\d+\.?\d*)s?)?/i.exec(value);
            let time = 0;
            const nagative = r[1];
            if (r[2]) {
                time += parseInt(r[2]) * 24 * 60 * 60 * 1000;
            } else if (r[3]) {
                time += parseInt(r[3]) * 60 * 60 * 1000;
            } else if (r[4]) {
                time += parseInt(r[4]) * 60 * 1000;
            } else if (r[5]) {
                time += Number(r[5]) * 1000;
            }
            return (nagative ? -1 : 1) * time;
        case 'bool':
            value = value && value.toLowerCase();
            return value == 'yes' || value == 'true' || value == '是' || value == 'y' || value == '1';
        case 'string':
            return value;
        default:
            let m, t;
            // int[]
            m = /(.*)\[\]/.exec(type);
            if (m) {
                if (value == '') return [];
                t = m[1];
                // value = value.replace(/，/g, ',');
                return value.split(',').map(function (v) {
                    return parseValue(v, t);
                });
            }
            // default string;
            return value;
    }
}

function checkArrayHasExplicitType(worksheet, startRow, startCol) {
    const typeNames = [
        'number',
        'int',
        'float',
        'time',
        'bool',
        'string',
    ]

    const type = worksheet.getRow(startRow + 1).getCell(startCol).text.trim()
    for (let i = 0; i < typeNames.length; i++) {
        if (type.indexOf(typeNames[i]) !== -1) {
            return true
        }
    }

    return false
}

function getColType(worksheet, col, dataStartRow) {
    const rowCount = worksheet.rowCount
    for (let i = dataStartRow; i <= rowCount; i++) {
        const text = worksheet.getRow(i).getCell(col).text
        if (text !== '') {
            return inferCellType(text)
        }
    }

    return null
}

function inferCellType(value) {
    if (value === '') return null

    if (value.indexOf(',') !== -1) {
        const arr = value.split(',')
        if (arr && arr[0]) return inferCellType(arr[0]) + '[]'
    }

    const numberValue = Number(value)
    if (!isNaN(numberValue)) {
        if (parseInt(value) === numberValue) {
            return 'int'
        } else {
            return 'float'
        }
    }

    const boolStrings = [
        'yes',
        'y',
        'true',
        '是',
        '1',
        'no',
        'n',
        'false',
        '否',
        '0',
    ]
    if (boolStrings.indexOf(value) !== -1) return 'bool'

    return 'string'
}

module.exports = {
    parseExcel: parseExcel,
    parseExcels: parseExcels,
};