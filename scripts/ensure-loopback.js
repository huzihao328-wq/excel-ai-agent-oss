/*
 * 静默允许 Edge WebView 访问 localhost，避免 office-addin-debugging
 * 弹出 “Allow localhost loopback for Microsoft Edge WebView? (Y/n)”。
 * 已放行则跳过；失败不阻断启动（与官方工具行为一致）。
 */
const { execFileSync } = require("child_process");

const EDGE_WEBVIEW = "Microsoft.win32webviewhost_cw5n1h2txyewy";

function listExemptions() {
  try {
    return execFileSync("CheckNetIsolation.exe", ["LoopbackExempt", "-s"], {
      encoding: "utf8",
      windowsHide: true,
    });
  } catch {
    return "";
  }
}

function alreadyAllowed(name) {
  return new RegExp(`Name:\\s*${name}`, "i").test(listExemptions());
}

if (process.platform !== "win32") process.exit(0);

if (alreadyAllowed(EDGE_WEBVIEW)) {
  console.log("localhost 回环已允许，跳过确认。");
  process.exit(0);
}

try {
  execFileSync("CheckNetIsolation.exe", ["LoopbackExempt", "-a", `-n=${EDGE_WEBVIEW}`], {
    encoding: "utf8",
    windowsHide: true,
  });
  console.log("已自动允许 Edge WebView 访问 localhost。");
} catch (e) {
  const detail = (e.stderr || e.stdout || e.message || "").toString().trim();
  console.warn("无法自动允许 localhost 回环，将继续启动。若插件打不开 localhost，请用管理员命令提示符执行：");
  console.warn(`  CheckNetIsolation.exe LoopbackExempt -a -n=${EDGE_WEBVIEW}`);
  if (detail) console.warn(detail);
}
