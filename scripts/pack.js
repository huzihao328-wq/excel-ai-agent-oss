/*
 * 默认打「内网离线安装包」：含 dist、本地 office.js、runtime/node.exe。
 * 目标机不需要访问 npm，也不需要先 npm install。
 *
 *   node scripts/pack.js          离线安装包
 *   node scripts/pack.js --src    仅源码（给能上网的开发机）
 */
const { execFileSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const OUT_DIR = path.join(ROOT, "pack");
const PKG = require(path.join(ROOT, "package.json"));
const SRC_ONLY = process.argv.includes("--src");

const SRC_INCLUDE = [
  "LICENSE",
  "package.json",
  "package-lock.json",
  "tsconfig.json",
  "webpack.config.js",
  "manifest.xml",
  "README.md",
  "MOTION.md",
  "AGENTS.md",
  "INSTALL.AGENT.md",
  "install.cmd",
  "uninstall.cmd",
  ".gitignore",
  "src",
  "scripts",
  "tests",
];

const OFFLINE_INCLUDE = [
  "LICENSE",
  "manifest.xml",
  "README.md",
  "AGENTS.md",
  "INSTALL.AGENT.md",
  "install.cmd",
  "uninstall.cmd",
  "scripts",
  "dist",
];

const SKIP_DIR = new Set(["node_modules", "pack", ".git", ".zcode", ".cache", "__pycache__", "test-results", "playwright-report"]);
const SKIP_FILE = new Set(["ExcelAIAgent.xlsx", ".DS_Store", "Thumbs.db", "desktop.ini", "motion-performance.json"]);

function shouldSkip(rel, isDir) {
  const base = path.basename(rel);
  if (isDir && SKIP_DIR.has(base)) return true;
  if (!isDir && SKIP_FILE.has(base)) return true;
  const normalized = rel.replace(/\\/g, "/");
  if (!isDir && (/^\.env(?:\.|$)/i.test(base) || /\.(?:log|tmp|bak|swp|pem|key|pfx|p12)$/i.test(base))) return true;
  if (!isDir && normalized.startsWith("tests/fixtures/") && normalized !== "tests/fixtures/ground-truth.json") return true;
  if (!isDir && normalized.startsWith("tests/") && /\.(?:xlsx?|xlsm|png|jpe?g|webp|gif|mp4|webm|zip)$/i.test(base)) return true;
  return false;
}

function copyTree(src, dest, rel) {
  const stat = fs.statSync(src);
  if (stat.isDirectory()) {
    if (rel && shouldSkip(rel, true)) return;
    fs.mkdirSync(dest, { recursive: true });
    for (const name of fs.readdirSync(src)) {
      copyTree(path.join(src, name), path.join(dest, name), rel ? path.join(rel, name) : name);
    }
    return;
  }
  if (shouldSkip(rel, false)) return;
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
}

function zipWithTar(stageDir, zipPath) {
  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
  execFileSync("tar", ["-a", "-cf", zipPath, "-C", stageDir, "."], { stdio: "inherit" });
}

function zipWithPowerShell(stageDir, zipPath) {
  if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
  const ps = `
$ErrorActionPreference = 'Stop'
Compress-Archive -Path (Join-Path '${stageDir.replace(/'/g, "''")}' '*') -DestinationPath '${zipPath.replace(/'/g, "''")}' -Force
`;
  execFileSync("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps], {
    stdio: "inherit",
    windowsHide: true,
  });
}

function run(script, args = []) {
  execFileSync(process.execPath, [path.join(__dirname, script), ...args], { cwd: ROOT, stdio: "inherit" });
}

function buildOfflineArtifacts() {
  execFileSync(process.execPath, [path.join(__dirname, "gen-icons.js")], { cwd: ROOT, stdio: "inherit" });
  run("vendor-office-js.js");
  const webpackCli = require.resolve("webpack-cli/bin/cli.js");
  execFileSync(process.execPath, [webpackCli, "--mode", "production"], { cwd: ROOT, stdio: "inherit" });
  if (!fs.existsSync(path.join(ROOT, "dist", "taskpane.html"))) {
    throw new Error("build 后仍没有 dist/taskpane.html");
  }
  if (!fs.existsSync(path.join(ROOT, "dist", "office-js", "office.js"))) {
    throw new Error("dist/office-js/office.js 缺失，任务窗格在内网会白屏");
  }
}

function copyBundledNode(inner) {
  const destDir = path.join(inner, "runtime");
  fs.mkdirSync(destDir, { recursive: true });
  const dest = path.join(destDir, "node.exe");
  fs.copyFileSync(process.execPath, dest);
  console.log("已带上 runtime/node.exe:", dest);
}

function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  if (!SRC_ONLY) buildOfflineArtifacts();
  else execFileSync(process.execPath, [path.join(__dirname, "gen-icons.js")], { cwd: ROOT, stdio: "inherit" });

  const stage = fs.mkdtempSync(path.join(os.tmpdir(), "excel-ai-agent-pack-"));
  const inner = path.join(stage, "excel-ai-agent");
  fs.mkdirSync(inner);

  const include = SRC_ONLY ? SRC_INCLUDE : OFFLINE_INCLUDE;
  for (const item of include) {
    const src = path.join(ROOT, item);
    if (!fs.existsSync(src)) {
      console.warn("跳过缺失项:", item);
      continue;
    }
    copyTree(src, path.join(inner, item), item);
  }

  if (SRC_ONLY) {
    const assetsSrc = path.join(ROOT, "assets");
    if (fs.existsSync(assetsSrc)) copyTree(assetsSrc, path.join(inner, "assets"), "assets");
    // 源码包不带第三方依赖；npm install 后，构建脚本会生成 vendor/office-js。
  } else {
    copyBundledNode(inner);
  }

  const zipName = SRC_ONLY ? `excel-ai-agent-${PKG.version}-src.zip` : `excel-ai-agent-${PKG.version}.zip`;
  const zipPath = path.join(OUT_DIR, zipName);
  try {
    zipWithTar(stage, zipPath);
  } catch (e) {
    console.warn("tar 打包失败，改用 Compress-Archive:", e.message);
    zipWithPowerShell(stage, zipPath);
  }
  fs.rmSync(stage, { recursive: true, force: true });

  const sizeMb = (fs.statSync(zipPath).size / 1024 / 1024).toFixed(2);
  console.log("");
  console.log(`已生成: ${zipPath}（${sizeMb} MB）`);
  if (SRC_ONLY) {
    console.log("这是源码包，目标机需要能访问 npm。内网请用默认的 npm run pack。");
  } else {
    console.log("内网安装：解压后双击 install.cmd（不需要 npm、不需要外网）。");
    console.log("升级已有机器：不要先卸载；装完彻底退出 Excel，再打开一次即可继承记忆和技能。");
  }
}

main();
