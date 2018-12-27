var fs = require('fs');
var path = require('path');
var parseExcel = require('./src/parse.js').parseExcel;
var parseExcels = require('./src/parse.js').parseExcels;
var render = require('./src/render.js');

function parse(filePath, outputDir, options) {
    if (!fs.existsSync(filePath)) {
        throw new Error('file: ' + filePath + ' not found');
    }
    options = options || {};
    if (isDirectory(filePath)) {
        var names = [];
        if (!outputDir) {
            throw new Error('need outputDir');
        }
        var mergeToOne = options.mergeToOne;
        var outputName = options.outputName;
        var fileNames = fs.readdirSync(filePath);
        var filePaths = [];
        fileNames.forEach(function(fileName) {
            names.push(fileName.split('.')[0]);
            filePaths.push(path.join(filePath, fileName));
        })
        parseExcels(filePaths, function(err, datas) {
            if (mergeToOne) {
                if (!outputName) {
                    throw new Error('need outputName');
                }
                var jsonData = {};
                names.forEach(function(name, idx) {
                    jsonData[name] = datas[idx];
                })
                render(path.join(outputDir, outputName + '.json'), jsonData);
            } else {
                names.forEach(function(name, idx) {
                    var jsonData = datas[idx];
                    render(path.join(outputDir, name + '.json'), jsonData);
                })
            }
            console.log('||||||||||||||||||||', err, datas);
        })
    } else {
        parseExcel(filePath, function(err, jsonData) {
            outputName = outputName || path.parse(filePath).name;
            render(path.join(outputDir, outputName + '.json'), jsonData);
        });
    }
}

function isDirectory(filePath) {
    var stat = fs.lstatSync(filePath);
    return stat.isDirectory();
}

module.exports = parse;