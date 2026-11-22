/**
 * 本机用户档案：记忆、技能、设置、自定义指令。
 * 写在 %LOCALAPPDATA%\ExcelAIAgent\userdata\，与插件源码/dist 分离，
 * 换版本、重装、清 WebView 缓存后仍能继承。
 *
 * 只响应来自本机的 GET/PUT /api/user-data。
 */
const fs = require("fs");
const os = require("os");
const path = require("path");

const MAX_BODY = 2 * 1024 * 1024;
const SCHEMA_VERSION = 1;

function dataDir() {
  const root = process.env.LOCALAPPDATA || process.env.HOME || os.homedir();
  return path.join(root, "ExcelAIAgent", "userdata");
}

function profilePath() {
  return path.join(dataDir(), "profile.json");
}

/**
 * 浏览器里只允许本插件页面访问：带了别的 Origin，或 Sec-Fetch-Site 表明跨站，一律拒绝。
 * 老版 Office WebView 不发这两个头，所以缺失时放行（仍受 loopback 限制）。
 */
function isAllowedCaller(req) {
  const origin = req.headers.origin;
  if (origin && origin !== "https://localhost:3000") return false;
  const site = req.headers["sec-fetch-site"];
  if (site && site !== "same-origin" && site !== "none") return false;
  return true;
}

function isLoopback(req) {
  const ip = (req.socket && req.socket.remoteAddress) || "";
  return ip === "127.0.0.1" || ip === "::1" || ip === "::ffff:127.0.0.1";
}

function pathnameOf(req) {
  const raw = req.url || "/";
  try {
    return decodeURIComponent(String(raw).split("?")[0]);
  } catch {
    return String(raw).split("?")[0];
  }
}

function send(res, status, body, contentType) {
  if (res.headersSent) return;
  const payload = typeof body === "string" ? body : JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": contentType || "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Origin": "https://localhost:3000",
    "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  });
  res.end(payload);
}

function readProfile() {
  const file = profilePath();
  if (!fs.existsSync(file)) return null;
  const raw = fs.readFileSync(file, "utf8");
  const parsed = JSON.parse(raw);
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return null;
  return parsed;
}

function writeProfile(obj) {
  const dir = dataDir();
  fs.mkdirSync(dir, { recursive: true });
  const tmp = path.join(dir, "profile.json.tmp");
  const file = profilePath();
  fs.writeFileSync(tmp, JSON.stringify(obj), "utf8");
  fs.renameSync(tmp, file);
}

function sanitizeProfile(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) return null;
  const out = {
    schemaVersion: SCHEMA_VERSION,
    updatedAt: typeof input.updatedAt === "number" ? input.updatedAt : Date.now(),
  };
  if (Array.isArray(input.skills)) out.skills = input.skills;
  if (Array.isArray(input.memories)) out.memories = input.memories;
  if (input.settings && typeof input.settings === "object" && !Array.isArray(input.settings)) {
    // API Key 不落盘（旧版本客户端仍可能发来，这里兜底剔除）
    const { apiKey: _omit, ...settings } = input.settings;
    out.settings = settings;
  }
  if (Array.isArray(input.templates)) out.templates = input.templates;
  return out;
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    const onData = (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY) {
        req.off("data", onData);
        reject(new Error("too large"));
        try {
          req.destroy();
        } catch {
          /* ignore */
        }
        return;
      }
      chunks.push(chunk);
    };
    req.on("data", onData);
    req.once("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.once("error", reject);
  });
}

/**
 * @returns {boolean} 是否已接管该请求
 */
function handleUserDataRequest(req, res) {
  if (pathnameOf(req) !== "/api/user-data") return false;

  if (!isLoopback(req) || !isAllowedCaller(req)) {
    send(res, 403, { error: "forbidden" });
    return true;
  }

  if (req.method === "OPTIONS") {
    send(res, 204, "");
    return true;
  }

  if (req.method === "GET") {
    try {
      const profile = readProfile();
      if (!profile) {
        send(res, 404, { error: "empty" });
        return true;
      }
      send(res, 200, profile);
    } catch (e) {
      send(res, 500, { error: e.message || String(e) });
    }
    return true;
  }

  if (req.method === "PUT") {
    readBody(req)
      .then((text) => {
        let parsed;
        try {
          parsed = JSON.parse(text || "{}");
        } catch {
          send(res, 400, { error: "invalid json" });
          return;
        }
        const profile = sanitizeProfile(parsed);
        if (!profile) {
          send(res, 400, { error: "invalid profile" });
          return;
        }
        writeProfile(profile);
        send(res, 200, { ok: true, updatedAt: profile.updatedAt });
      })
      .catch((e) => {
        const status = e && e.message === "too large" ? 413 : 500;
        send(res, status, { error: e.message || String(e) });
      });
    return true;
  }

  send(res, 405, { error: "method not allowed" });
  return true;
}

module.exports = { handleUserDataRequest, dataDir, profilePath };
