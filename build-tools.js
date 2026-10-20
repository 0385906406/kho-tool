// Tạo tools.js cho bản Kho Tool trên claude.ai: danh sách tool + file nhúng base64.
// Danh sách tool nằm ở tools.config.js. Chạy: node build-tools.js
const fs = require('fs'), path = require('path');
const here = __dirname;
const TOOLS = require('./tools.config.js').map(t => Object.assign({}, t, {
  files: t.files.filter(f => !f.webOnly && !f.xlsmOnly).map(f => ({
    label: f.label, name: f.name, kind: f.kind,
    size: fs.statSync(path.join(here, f.src)).size,
    b64: fs.readFileSync(path.join(here, f.src)).toString('base64'),
  })),
}));
fs.writeFileSync(path.join(here, 'tools.js'), '// Sinh tự động bởi build-tools.js — không sửa tay\nwindow.TOOLS = ' + JSON.stringify(TOOLS) + ';\n');
console.log('tools.js', fs.statSync(path.join(here, 'tools.js')).size, 'bytes');
