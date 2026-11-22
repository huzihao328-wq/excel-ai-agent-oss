/*
 * 把静态服务器设置为 Windows 开机自启动（写入用户"启动"文件夹的 VBS 脚本，隐藏窗口运行），
 * 并立即在后台启动一次。
 *
 * 用法:
 *   node scripts/setup-autostart.js         安装自启动并立即启动
 *   node scripts/setup-autostart.js remove  移除自启动
 */
const fs = require("fs");
const path = require("path");
const { execFileSync, spawn } = require("child_process");
const { projectDir, resolveNode } = require("./runtime-node");

const VBS_NAME = "excel-ai-agent-server.vbs";
const startupDir = path.join(
  process.env.APPDATA || "",
  "Microsoft",
  "Windows",
  "Start Menu",
  "Programs",
  "Startup"
);
const vbsPath = path.join(startupDir, VBS_NAME);
const servePath = path.join(__dirname, "serve-dist.js");
const nodePath = resolveNode();

function stopServeProcess() {
  try {
    execFileSync(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "Get-CimInstance Win32_Process | Where-Object { $_.Name -match '^(node|node.exe)$' -and $_.CommandLine -like '*serve-dist.js*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; Start-Sleep -Milliseconds 400",
      ],
      { windowsHide: true, timeout: 20000 }
    );
  } catch {
    /* ignore */
  }
}

if (process.argv[2] === "remove") {
  stopServeProcess();
  if (fs.existsSync(vbsPath)) {
    fs.unlinkSync(vbsPath);
    console.log("已移除开机自启动:", vbsPath);
  } else {
    console.log("未找到自启动项，无需移除。");
  }
  process.exit(0);
}

if (!fs.existsSync(startupDir)) {
  console.error("未找到启动文件夹:", startupDir);
  process.exit(1);
}

// VBS 里的路径用双引号包裹，内部双引号翻倍转义
const q = (s) => `""${s}""`;
const vbs = [
  `Set WshShell = CreateObject("WScript.Shell")`,
  `WshShell.CurrentDirectory = "${projectDir}"`,
  `WshShell.Run "${q(nodePath)} ${q(servePath)}", 0, False`,
  "",
].join("\r\n");

fs.writeFileSync(vbsPath, vbs, "utf-8");
console.log("已写入开机自启动:", vbsPath);

stopServeProcess();

const child = spawn(nodePath, [servePath], {
  detached: true,
  stdio: "ignore",
  windowsHide: true,
  cwd: projectDir,
});
child.unref();
console.log("静态服务已在后台启动: https://localhost:3000");
console.log("下次开机登录后会自动启动，无需手动运行任何命令。");
