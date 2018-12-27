var fs = require('fs');
function render(filePath, jsonData) {
    fs.writeFileSync(filePath, JSON.stringify(jsonData));
}
module.exports = render;