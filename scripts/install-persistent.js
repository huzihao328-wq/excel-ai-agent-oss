/**
 * 把插件装成 Excel 重启后仍自动加载：
 * 1. 注册开发者清单（WEF\Developer）
 * 2. 建立本机共享目录并写入 TrustedCatalogs
 * 3. 开机自动启动 https://localhost:3000 静态服务
 *    （不再往 XLSTART 写 Book.xltx，避免 Excel 启动时报格式无效）
 *
 *   node scripts/install-persistent.js
 *   node scripts/install-persistent.js remove
 */
const { execFileSync, spawnSync } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");

const ADDIN_ID = "7c1f5f2e-8a4b-4d3c-9e6f-2b0a1c9d8e7f";
const SHARE_NAME = "ExcelAIAgentCatalog";
const CATALOG_GUID = `{${ADDIN_ID}}`;
const XLSTART_NAME = "Book.xltx";
const OLD_XLAM = "ExcelAIAgentWeb.xlam";

const projectDir = path.join(__dirname, "..");
const manifestSrc = path.join(projectDir, "manifest.xml");
const catalogDir = path.join(process.env.LOCALAPPDATA || "", "ExcelAIAgent", "catalog");
const xlstartDir = path.join(process.env.APPDATA || "", "Microsoft", "Excel", "XLSTART");
const xlstartPath = path.join(xlstartDir, XLSTART_NAME);
const shareUnc = `\\\\${os.hostname()}\\${SHARE_NAME}`;

function ps(command, opts = {}) {
  return execFileSync(
    "powershell.exe",
    ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
    { encoding: "utf8", windowsHide: true, ...opts }
  );
}

function psElevated(command) {
  const encoded = Buffer.from(command, "utf16le").toString("base64");
  const r = spawnSync(
    "powershell.exe",
    [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-Command",
      `Start-Process powershell.exe -Verb RunAs -Wait -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand','${encoded}'`,
    ],
    { encoding: "utf8", windowsHide: true }
  );
  if (r.status !== 0) {
    throw new Error((r.stderr || r.stdout || "elevate failed").trim());
  }
}

function remove() {
  require("child_process").execFileSync(process.execPath, [path.join(__dirname, "setup-autostart.js"), "remove"], {
    stdio: "inherit",
    cwd: projectDir,
  });

  try {
    unregisterDeveloper();
  } catch {
    /* already unregistered */
  }

  for (const name of [XLSTART_NAME, OLD_XLAM]) {
    const p = path.join(xlstartDir, name);
    if (fs.existsSync(p)) {
      fs.unlinkSync(p);
      console.log("已移除 XLSTART 文件:", p);
    }
  }

  const catKey = `HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\${CATALOG_GUID}`;
  ps(`if (Test-Path -LiteralPath '${catKey}') { Remove-Item -LiteralPath '${catKey}' -Recurse -Force }`);
  console.log("已移除 TrustedCatalogs 注册");

  try {
    psElevated(
      `if (Get-SmbShare -Name '${SHARE_NAME}' -ErrorAction SilentlyContinue) { Remove-SmbShare -Name '${SHARE_NAME}' -Force }`
    );
    console.log("已移除共享目录", SHARE_NAME);
  } catch (e) {
    console.warn("移除共享失败（可忽略）:", e.message);
  }
}

function ensureCatalogShare() {
  fs.mkdirSync(catalogDir, { recursive: true });
  fs.copyFileSync(manifestSrc, path.join(catalogDir, "manifest.xml"));
  console.log("已更新目录清单:", path.join(catalogDir, "manifest.xml"));

  const shareCmd = `
$ErrorActionPreference = 'Stop'
$name = '${SHARE_NAME}'
$path = '${catalogDir.replace(/'/g, "''")}'
$existing = Get-SmbShare -Name $name -ErrorAction SilentlyContinue
if ($existing -and $existing.Path -ne $path) {
  Remove-SmbShare -Name $name -Force
  $existing = $null
}
if (-not $existing) {
  New-SmbShare -Name $name -Path $path -ReadAccess 'Everyone' | Out-Null
}
icacls $path /grant '*S-1-1-0:(OI)(CI)R' /T | Out-Null
`.trim();

  try {
    psElevated(shareCmd);
    console.log("已创建/确认共享:", shareUnc);
  } catch (e) {
    console.warn("创建 SMB 共享失败，将只依赖开发者注册。", e.message);
  }

  const key = `HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\${CATALOG_GUID}`;
  ps(`
New-Item -Path '${key}' -Force | Out-Null
New-ItemProperty -Path '${key}' -Name Id -Value '${CATALOG_GUID}' -PropertyType String -Force | Out-Null
New-ItemProperty -Path '${key}' -Name Url -Value '${shareUnc.replace(/'/g, "''")}' -PropertyType String -Force | Out-Null
New-ItemProperty -Path '${key}' -Name Flags -Value 1 -PropertyType DWord -Force | Out-Null
`);
  console.log("已写入 TrustedCatalogs:", CATALOG_GUID);
}

function unregisterDeveloper() {
  const manifestAbs = path.resolve(manifestSrc).replace(/'/g, "''");
  ps(`
$key = 'HKCU:\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer'
if (Test-Path $key) {
  Remove-ItemProperty -LiteralPath $key -Name '${ADDIN_ID}' -ErrorAction SilentlyContinue
  Get-ItemProperty $key -ErrorAction SilentlyContinue | ForEach-Object {
    $_.PSObject.Properties | Where-Object { $_.Value -eq '${manifestAbs}' } | ForEach-Object {
      Remove-ItemProperty -LiteralPath $key -Name $_.Name -ErrorAction SilentlyContinue
    }
  }
}
$sub = Join-Path $key '${ADDIN_ID}'
if (Test-Path $sub) { Remove-Item -LiteralPath $sub -Recurse -Force }
`);
  console.log("已移除开发者清单注册");
}

function registerDeveloper() {
  const manifestAbs = path.resolve(manifestSrc).replace(/'/g, "''");
  ps(`
$key = 'HKCU:\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer'
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name '${ADDIN_ID}' -Value '${manifestAbs}' -PropertyType String -Force | Out-Null
$sub = Join-Path $key '${ADDIN_ID}'
New-Item -Path $sub -Force | Out-Null
New-ItemProperty -Path $sub -Name UseDirectDebugger -Value 0 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $sub -Name UseWebDebugger -Value 0 -PropertyType DWord -Force | Out-Null
Remove-ItemProperty -LiteralPath $sub -Name OpenDevTools -ErrorAction SilentlyContinue
`);
  console.log("已注册开发者清单（未访问 npm）:", path.resolve(manifestSrc));
}

function setupServer() {
  execFileSync(process.execPath, [path.join(__dirname, "setup-autostart.js")], {
    stdio: "inherit",
    cwd: projectDir,
  });
}

function main() {
  if (process.argv[2] === "remove") {
    remove();
    return;
  }
  if (!fs.existsSync(path.join(projectDir, "dist", "taskpane.html"))) {
    throw new Error("dist/ 不完整。离线包应已包含 dist/；开发机请先 npm run build");
  }
  try {
    execFileSync(process.execPath, [path.join(__dirname, "ensure-loopback.js")], {
      stdio: "inherit",
      cwd: projectDir,
    });
  } catch {
    /* 回环放行失败不阻断 */
  }
  registerDeveloper();
  ensureCatalogShare();
  // 不再往 XLSTART 写 Book.xltx：把 xlsx 改后缀当模板会被 Excel 拒收，
  // 每次启动都弹「文件格式或扩展名无效」。插件靠开发者注册 + 信任目录加载即可。
  for (const leftover of [xlstartPath, path.join(xlstartDir, OLD_XLAM)]) {
    if (fs.existsSync(leftover)) {
      fs.unlinkSync(leftover);
      console.log("已移除 XLSTART 残留:", leftover);
    }
  }
  setupServer();
  console.log("");
  console.log("已设置为自动加载。请完全退出 Excel 再重新打开。");
  console.log("之后每次开 Excel，开始选项卡都应有 AI Agent。");
}

main();
