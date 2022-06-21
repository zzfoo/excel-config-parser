var fs = require('fs');
var path = require('path');
var parseExcel = require('./src/parse.js').parseExcel;
var parseExcels = require('./src/parse.js').parseExcels;
var render = require('./src/render.js');
var config = require('./src/config.js');

// options = {
//     outputName: '',
//     pretty: false,
//     mergeToOne: false,
//     sheetAsStandaloneFile: false,
//     startRow: 1, (1-based)
//     invalidRowMark: '' // '#'
// }
function parse(filePath, outputDir, options, callback) {
    if (!fs.existsSync(filePath)) {
        throw new Error('file: ' + filePath + ' not found');
    }

    if (!outputDir) {
        throw new Error('need outputDir');
    }

    config.setOptions(options)
    options = config.getOptions()

    var mergeToOne = options.mergeToOne;
    var outputName = options.outputName;
    var sheetAsStandaloneFile = options.sheetAsStandaloneFile;

    var names = [];
    var filePaths = [];
    if (isDirectory(filePath)) {
        var fileNames = fs.readdirSync(filePath);
        fileNames.forEach(function (fileName) {
            var info = fileName.split('.');
            if (info[1] !== 'xlsx') {
                return;
            }
            if (fileName.indexOf('~') === 0) {
                return;
            }
            names.push(info[0]);
            filePaths.push(path.join(filePath, fileName));
        })
    } else {
        filePaths.push(filePath);
        outputName = outputName || path.parse(filePath).name;
        names = [outputName]
    }

    parseExcels(filePaths, function (err, datas) {
        if (mergeToOne) {
            if (!outputName) {
                throw new Error('need outputName');
            }
            var jsonData = {};
            names.forEach(function (name, idx) {
                jsonData[name] = datas[idx];
            })
            render(path.join(outputDir, outputName + '.json'), jsonData);
        } else if (sheetAsStandaloneFile) {
            datas.forEach((data) => {
                for (var key in data) {
                    render(path.join(outputDir, key + '.json'), data[key]);
                }
            })
        } else {
            names.forEach(function (name, idx) {
                var jsonData = datas[idx];
                render(path.join(outputDir, name + '.json'), jsonData);
            })
        }

        callback && callback()
    })
}

function isDirectory(filePath) {
    var stat = fs.lstatSync(filePath);
    return stat.isDirectory();
}

module.exports = parse;