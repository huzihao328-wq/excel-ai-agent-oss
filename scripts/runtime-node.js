const fs = require("fs");
const path = require("path");

const projectDir = path.join(__dirname, "..");

/** 优先用离线包自带的 node.exe，避免目标机安装 Node / 访问 npm */
function resolveNode() {
  const bundled = path.join(projectDir, "runtime", "node.exe");
  if (process.platform === "win32" && fs.existsSync(bundled)) return bundled;
  return process.execPath;
}

module.exports = { projectDir, resolveNode };
