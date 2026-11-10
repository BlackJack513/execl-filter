/*
 * 把模板、样式、ExcelJS 与业务脚本打包成单个 HTML。
 * 用法：node build.js
 */
const fs = require('fs');
const path = require('path');

const root = __dirname;
const templatePath = path.join(root, 'src', 'index.template.html');
const outputDir = path.join(root, 'dist');
const outputPath = path.join(outputDir, '打卡无效记录筛查.html');

function read(file) {
  return fs.readFileSync(file, 'utf8');
}

function inject(html, placeholder, content, label) {
  const count = html.split(placeholder).length - 1;
  if (count !== 1) {
    throw new Error(`占位符 ${placeholder} 应出现 1 次，实际 ${count} 次`);
  }
  if (/<\/script/i.test(content)) {
    throw new Error(`${label} 含有 </script，不能安全内联`);
  }
  return html.replace(placeholder, () => content);
}

const template = read(templatePath);
const style = read(path.join(root, 'src', 'style.css'));
const core = read(path.join(root, 'src', 'core.js'));
const app = read(path.join(root, 'src', 'app.js'));
const exceljs = read(path.join(root, 'node_modules', 'exceljs', 'dist', 'exceljs.min.js'));

let html = template;
html = inject(html, '/*__STYLE__*/', style, 'style.css');
html = inject(html, '/*__EXCELJS__*/', exceljs, 'exceljs.min.js');
html = inject(html, '/*__CORE__*/', core, 'core.js');
html = inject(html, '/*__APP__*/', app, 'app.js');

if (html.includes('/*__')) {
  throw new Error('仍有未替换的占位符');
}

fs.mkdirSync(outputDir, { recursive: true });
fs.writeFileSync(outputPath, html, 'utf8');

const size = Buffer.byteLength(html, 'utf8');
console.log(`已生成：${outputPath}`);
console.log(`大小：${(size / 1024 / 1024).toFixed(2)} MB（${size} 字节）`);
