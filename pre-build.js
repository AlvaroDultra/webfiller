const fs = require('fs');
const path = './node_modules/.bin/webpack';

fs.chmodSync(path, '755');
