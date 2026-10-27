// 檢查 vendor/ 內套件未被竄改（與 SHA256SUMS 比對），並確認版本與 package.json 一致
const fs = require('fs'), path = require('path'), crypto = require('crypto');
const root = path.join(__dirname, '..');
let ok = true;
for (const line of fs.readFileSync(path.join(root, 'vendor', 'SHA256SUMS'), 'utf8').trim().split('\n')) {
  const [hash, name] = line.split(/\s+/);
  const h = crypto.createHash('sha256').update(fs.readFileSync(path.join(root, 'vendor', name))).digest('hex');
  console.log((h === hash ? 'OK   ' : 'FAIL ') + name); if (h !== hash) ok = false;
}
const pkg = require(path.join(root, 'package.json'));
const xl = fs.readFileSync(path.join(root, 'vendor', 'xlsx.full.min.js'), 'utf8');
if (!xl.includes(pkg.devDependencies.xlsx)) { console.log('FAIL xlsx 版本與 package.json 不符'); ok = false; }
const jz = fs.readFileSync(path.join(root, 'vendor', 'jszip.min.js'), 'utf8');
if (!jz.includes('v' + pkg.devDependencies.jszip)) { console.log('FAIL jszip 版本與 package.json 不符'); ok = false; }
process.exit(ok ? 0 : 1);
