var addon = require('bindings')('office');

// console.log(addon.lok_init("/usr/lib/libreoffice/program/")); 

const office = new addon.Office();
office.load("/usr/lib/libreoffice/program/")
office.load_document("./test/index.html")
office.destroy();
