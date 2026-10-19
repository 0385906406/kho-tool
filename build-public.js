// Tạo bản web độc lập trong thư mục public/ (không cần claude.ai, không cần đăng nhập).
// Đưa nguyên thư mục public/ lên Netlify / GitHub Pages / ổ chung là dùng được.
// Danh sách tool nằm ở tools.config.js. Chạy: node build-public.js
const fs = require('fs'), path = require('path');
const here = __dirname, out = path.join(here, 'public'), filesDir = path.join(out, 'files');
fs.mkdirSync(filesDir, { recursive: true });
const TOOLS = require('./tools.config.js').map(t => Object.assign({}, t, {
  files: t.files.map(f => {
    const dir = path.join(filesDir, t.id);
    fs.mkdirSync(dir, { recursive: true });
    fs.copyFileSync(path.join(here, f.src), path.join(dir, f.name));
    return { label: f.label, name: f.name, kind: f.kind, size: fs.statSync(path.join(here, f.src)).size, url: 'files/' + t.id + '/' + encodeURIComponent(f.name) };
  }),
}));
fs.writeFileSync(path.join(out, 'tools.js'), '// Sinh tự động bởi build-public.js — không sửa tay\nwindow.TOOLS = ' + JSON.stringify(TOOLS) + ';\n');
fs.copyFileSync(path.join(here, 'public-index.html'), path.join(out, 'index.html'));
console.log('public/ ready:', TOOLS.length, 'tool,', TOOLS.reduce((n, t) => n + t.files.length, 0), 'file');
