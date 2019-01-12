var exceljs = require('exceljs');
var async = require('async');

function parseExcel(filePath, callback) {
    console.log('===================================');
    console.log('===================================');
    console.log('===================================');
    console.log('parsing excel: ', filePath);
    var data = {};
    var workbook = new exceljs.Workbook();
    workbook.xlsx.readFile(filePath)
        .then(function() {
            var worksheets = workbook.worksheets;
            worksheets.forEach(function(sheet) {
                var name = sheet.name;
                console.log('===================================');
                console.log('parsing sheet: ', name);
                var sheetData = parseSheet(sheet);
                console.log('sheetData: ', sheetData);
                data[name] = sheetData;
            })
            console.log('data: ', data);
            callback(null, data);
        })
        .catch(function(e) {
            callback(e);
        })
}

function parseExcels(filePaths, callback) {
    async.map(filePaths, parseExcel, function(err, datas) {
        callback(err, datas);
    })
}

// function parseExcels(filePaths, callback) {
//     var map = {};
//     var datas = [];
//     filePaths.forEach(function(filePath, idx) {
//         map[filePath] = idx;
//     })
//     async.forEachOf(map, function(idx, filePath, cb) {
//         parseExcel(filePath, function(err, data) {
//             if (err) {
//                 cb(err);
//                 return;
//             }
//             datas[idx] = data;
//             cb(null, data);
//         });
//     }, function(err, datas) {
//         callback(err, datas);
//     })
// }

function parseSheet(worksheet) {
    var firstRow = worksheet.getRow(1);
    if (firstRow.getCell(1).value === 'key' && firstRow.getCell(2).value === 'value' && firstRow.getCell(3).value === 'type') {
        console.log('type: map');
        return parseMap(worksheet);
    }
    console.log('type: array');
    return parseArray(worksheet);
}

function parseMap(worksheet) {
    var data = {};
    var key;
    var value;
    var type;
    worksheet.eachRow(function(row, rowNumber) {
        if (rowNumber <= 1) {
            return;
        }
        key = row.getCell(1).value;
        value = row.getCell(2).value;
        type = row.getCell(3).value;
        if (!key || !type) {
            throw new Error('worksheet: ' + worksheet.name + ' ' + rowNumber + ' key or type is empty!');
        }
        try {
            data[key] = parseValue(value, type);
        } catch(e) {
            console.error('worksheet: ' + worksheet.name + ' row: ' + rowNumber + ' parse failed! ');
            console.error('key: ' + key + ' , value: ' + value + ' , type: ' + type);
            throw e;
        }
    })
    return data;
}

function parseArray(worksheet) {
    var data = [];

    var keys = [];
    worksheet.getRow(1).eachCell(function(cell, cellNumber) {
        var key;
        key = cell.value;
        if (key.trim) {
            key = key.trim();
        }
        if (!key) {
            throw new Error('worksheet: ' + worksheet.name + ' key is empty!');
        }
        keys.push(key);
    })

    var types = [];
    worksheet.getRow(2).eachCell(function(cell, cellNumber) {
        var type;
        type = cell.value.trim();
        if (type.trim) {
            type = type.trim();
        }
        if (!type) {
            throw new Error('worksheet: ' + worksheet.name + ' type is empty!');
        }
        types.push(type);
    })

    var item;
    worksheet.eachRow(function(row, rowNumber) {
        if (rowNumber <= 2) {
            return;
        }

        item = {};
        keys.forEach(function(key, keyIndex) {
            var type = types[keyIndex];
            var value = row.getCell(keyIndex + 1).value;
            try {
                item[key] = parseValue(value, type);
            } catch(e) {
                console.log(e);
                console.error('worksheet: ' + worksheet.name + ' row: ' + rowNumber + ' parse failed! ');
                console.error('key: ' + key + ' , value: ' + value + ' , type: ' + type);
                throw e;
            }
        });
        data.push(item);
    })
    return data;
}

function parseValue(value, type) {
    if (typeof value === 'string') {
        value = value.trim();
    }
    if(value == 'N/A'||value == 'n/a') return undefined;
    if(value === 'null') return null;
    switch(type) {
      case 'int':
        if (value === null) return undefined;
        return parseInt(value);
      case 'num':
      case 'float':
      case 'number':
        if (value === null) return undefined;
        return Number(value);
      case 'time':
        if (value === null) return undefined;
        var m = /(-?)(?:(\d+)d)?(?:(\d+)h)?(?:(\d+)m)?(?:(\d+\.?\d*)s?)?/i.exec(value);
        var time = 0;
        var nagative = m[1];
        if(m[2]) {
            time += parseInt(m[2]) * 24 * 60 * 60 * 1000;
        } else if(m[3]) {
            time += parseInt(m[3]) * 60 * 60 * 1000;
        } else if(m[4]) {
            time += parseInt(m[4]) * 60 * 1000;
        } else if(m[5]) {
            time += Number(m[5]) * 1000;
        }
        return (nagative ? -1 : 1) * time;
      case 'bool':
        if (value === null) return undefined;
        value = value && value.toLowerCase();
        return value == 'yes' || value == 'true' || value == '是' || value == 'y' || value == '1';
      case 'string':
        return value;
      default:
        var m, t;
        // int[]
        m = /(.*)\[\]/.exec(type);
        if(m) {
            if(value == '') return [];
            t = m[1];
            value = value.replace(/，/g, ',');
            return value.split(',').map(function(v) {
                    return parseValue(v, t);
            });
        }
        // default string;
        return value;
    }
}

module.exports = {
    parseExcel: parseExcel,
    parseExcels: parseExcels,
};