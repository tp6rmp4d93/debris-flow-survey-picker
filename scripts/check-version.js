// 檢查 package.json、js/version.js、CHANGELOG.md 的版本號一致
const fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..');
const pkg = require(path.join(root, 'package.json')).version;
const js = (fs.readFileSync(path.join(root, 'js', 'version.js'), 'utf8').match(/APP_VERSION\s*=\s*'([^']+)'/) || [])[1];
const cl = (fs.readFileSync(path.join(root, 'CHANGELOG.md'), 'utf8').match(/^## \[(\d+\.\d+\.\d+)\]/m) || [])[1];
console.log({ 'package.json': pkg, 'js/version.js': js, 'CHANGELOG.md（最新）': cl });
if (pkg !== js || pkg !== cl) { console.error('版本號不一致，請同步更新三處'); process.exit(1); }
