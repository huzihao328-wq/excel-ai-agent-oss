// OpenCode Go does not allow browser CORS requests. Relay only its fixed API
// from our local taskpane; credentials are forwarded, never stored or logged.
const https = require("https");
const { randomUUID } = require("crypto");
const SESSION = randomUUID();
const PREFIX = "/api/opencode-go";

function handleOpenCodeRequest(req, res) {
  const route = (req.url || "").split("?")[0];
  if (!route.startsWith(PREFIX + "/")) return false;
  const fail = (status, message) => {
    if (res.destroyed) return;
    if (res.headersSent) { res.destroy(); return; }
    res.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" });
    res.end(JSON.stringify({ error: { message } }));
  };
  const ip = req.socket.remoteAddress;
  const origin = req.headers.origin;
  const sameOrigin = origin === "https://localhost:3000" ||
    (!origin && req.headers["sec-fetch-site"] === "same-origin");
  if (!["127.0.0.1", "::1", "::ffff:127.0.0.1"].includes(ip) || !sameOrigin) {
    fail(403, "Only the local Excel taskpane may use this endpoint.");
    return true;
  }
  const suffix = route.slice(PREFIX.length);
  if (!((req.method === "POST" && suffix === "/chat/completions") ||
    (req.method === "GET" && /^\/models(?:\/[a-zA-Z0-9_.%-]+)?$/.test(suffix)))) {
    fail(404, "Unsupported OpenCode endpoint.");
    return true;
  }
  if (!/^Bearer \S+$/i.test(req.headers.authorization || "")) {
    fail(401, "API Key is required.");
    return true;
  }
  const upstream = https.request("https://opencode.ai/zen/go/v1" + suffix, {
    method: req.method,
    headers: {
      Authorization: req.headers.authorization,
      "Content-Type": "application/json",
      Accept: req.headers.accept || "application/json",
      "User-Agent": "ExcelAIAgent/0.1.0",
      "x-opencode-session": SESSION,
    },
  }, (response) => {
    if (res.destroyed) { response.destroy(); return; }
    res.writeHead(response.statusCode || 502, {
      "Content-Type": response.headers["content-type"] || "application/json",
      "Cache-Control": "no-store",
    });
    res.flushHeaders();
    response.on("error", () => res.destroy());
    response.pipe(res);
  });
  upstream.setTimeout(120000, () => upstream.destroy(new Error("OpenCode request timed out.")));
  upstream.on("error", () => fail(502, "OpenCode connection failed or timed out. Please retry."));
  req.on("aborted", () => upstream.destroy());
  res.on("close", () => upstream.destroy());
  let bytes = 0;
  req.on("data", (chunk) => {
    bytes += chunk.length;
    if (bytes > 8 * 1024 * 1024) {
      req.unpipe(upstream);
      fail(413, "Request exceeds 8 MB.");
      upstream.destroy();
    }
  });
  req.pipe(upstream);
  return true;
}

module.exports = { handleOpenCodeRequest };
