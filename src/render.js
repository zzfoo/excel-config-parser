var fs = require('fs');
var config = require('./config.js')
function render(filePath, jsonData) {
    var options = config.getOptions()
    if (options.pretty) {
        fs.writeFileSync(filePath, JSON.stringify(jsonData, null, 4));
    } else {
        fs.writeFileSync(filePath, JSON.stringify(jsonData));
    }
}
module.exports = render;