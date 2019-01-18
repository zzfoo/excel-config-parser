var fs = require('fs');
function render(filePath, jsonData, options) {
    if (options.pretty) {
        fs.writeFileSync(filePath, JSON.stringify(jsonData, null, 4));
    } else {
        fs.writeFileSync(filePath, JSON.stringify(jsonData));
    }
}
module.exports = render;