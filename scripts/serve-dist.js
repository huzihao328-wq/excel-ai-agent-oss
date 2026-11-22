/*
 * 生产静态服务器：https://localhost:3000 提供 dist/。
 * 证书优先用本机自签（不依赖 npm）；若装了 office-addin-dev-certs 也可沿用。
 */
const https = require("https");
const fs = require("fs");
const path = require("path");
const { ensureLocalhostCert } = require("./localhost-certs");
const { handleUserDataRequest } = require("./userdata-store");
const { handleOpenCodeRequest } = require("./opencode-proxy");

const DIST = path.join(__dirname, "..", "dist");
const PORT = 3000;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "text/xml; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
};

function handler(req, res) {
  if (handleOpenCodeRequest(req, res)) return;
  if (handleUserDataRequest(req, res)) return;
  let urlPath;
  try {
    urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  } catch {
    // 畸形百分号编码（如 "/%"）：任何网页都能发这种请求，不能让它打崩常驻服务
    res.writeHead(400);
    res.end("Bad Request");
    return;
  }
  const filePath = path.normalize(path.join(DIST, urlPath === "/" ? "taskpane.html" : urlPath));
  const rel = path.relative(DIST, filePath);
  // 用 relative 判断而不是 startsWith：后者会放过同级的 dist-xxx 目录
  if (!rel || rel.startsWith("..") || path.isAbsolute(rel)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not Found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream",
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-cache",
    });
    res.end(data);
  });
}

async function loadCerts() {
  try {
    const devCerts = require("office-addin-dev-certs");
    return await devCerts.getHttpsServerOptions();
  } catch {
    return ensureLocalhostCert();
  }
}

function safeHandler(req, res) {
  try {
    handler(req, res);
  } catch (e) {
    console.error("请求处理出错:", e && e.stack ? e.stack : e);
    if (!res.headersSent) res.writeHead(500);
    res.end();
  }
}

// 常驻服务只在登录时由启动项拉起一次，没有守护进程；单个请求的意外异常不能让整个插件失效到下次登录。
process.on("uncaughtException", (e) => {
  console.error("未捕获异常（服务继续运行）:", e && e.stack ? e.stack : e);
});

(async () => {
  if (!fs.existsSync(path.join(DIST, "taskpane.html"))) {
    console.error("dist/ 不存在或不完整。离线包应已包含 dist/；开发机请先 npm run build");
    process.exit(1);
  }
  const certs = await loadCerts();
  const server = https.createServer(certs, safeHandler);
  server.on("error", (e) => {
    if (e.code === "EADDRINUSE") {
      console.log(`端口 ${PORT} 已被占用，无需重复启动。`);
      process.exit(0);
    }
    console.error("静态服务启动失败:", e && e.message ? e.message : e);
    process.exit(1);
  });
  server.listen(PORT, () => {
    console.log(`Excel AI Agent 静态服务已启动: https://localhost:${PORT}`);
  });
})();
