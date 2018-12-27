var path = require('path');
var parse = require('../index.js');
// parse(path.resolve(__dirname, './input/config.xlsx'), path.resolve(__dirname, './output/'));
// parse(path.resolve(__dirname, './input/'), path.resolve(__dirname, './output/'), {
//     mergeToOne: true,
//     outputName: 'config',
// });
parse(path.resolve(__dirname, './input/'), path.resolve(__dirname, './output/'), {
});