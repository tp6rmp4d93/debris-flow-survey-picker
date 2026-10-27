// 由 node_modules 複製前端套件到 vendor/ 並更新 SHA256SUMS
// 用法：npm ci && npm run vendor:update
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const root = path.join(__dirname, '..');
const files = [
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'vendor/xlsx.full.min.js'],
  ['node_modules/xlsx/LICENSE', 'vendor/LICENSE-xlsx.txt'],
  ['node_modules/jszip/dist/jszip.min.js', 'vendor/jszip.min.js'],
  ['node_modules/jszip/LICENSE.markdown', 'vendor/LICENSE-jszip.md']
];
for (const [a, b] of files) fs.copyFileSync(path.join(root, a), path.join(root, b));
const sums = ['jszip.min.js', 'xlsx.full.min.js'].map(f =>
  crypto.createHash('sha256').update(fs.readFileSync(path.join(root, 'vendor', f))).digest('hex') + '  ' + f).join('\n') + '\n';
fs.writeFileSync(path.join(root, 'vendor', 'SHA256SUMS'), sums);
console.log(sums);
