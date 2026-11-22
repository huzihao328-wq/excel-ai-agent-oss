/*
 * 把 Office.js 拷到 vendor/office-js，构建后随 dist 分发，任务窗格不再访问微软 CDN。
 */
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const dest = path.join(ROOT, "vendor", "office-js");
const src = path.join(ROOT, "node_modules", "@microsoft", "office-js", "dist");

function copyDir(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const name of fs.readdirSync(from)) {
    const a = path.join(from, name);
    const b = path.join(to, name);
    if (fs.statSync(a).isDirectory()) copyDir(a, b);
    else fs.copyFileSync(a, b);
  }
}

if (fs.existsSync(path.join(dest, "office.js"))) {
  console.log("已存在 vendor/office-js，跳过。");
  process.exit(0);
}

if (!fs.existsSync(path.join(src, "office.js"))) {
  console.error("未找到 @microsoft/office-js。请在能访问 npm 的机器上执行: npm install");
  process.exit(1);
}

copyDir(src, dest);
console.log("已写入 vendor/office-js（离线打包用，不再请求 appsforoffice.microsoft.com）");
