const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = Number(process.env.PORT || 3001);
const MIMO_BASE = "https://platform.xiaomimimo.com";
const POLL_INTERVAL = 10 * 1000; // 10 seconds
const STORE_PATH = path.join(__dirname, "store.json");
const ENV_PATH = path.join(__dirname, "..", ".env");

// --- Cookie management ---
function loadCookies() {
  // Priority: store.json > .env
  try {
    const store = JSON.parse(fs.readFileSync(STORE_PATH, "utf-8"));
    if (store.cookies) return store.cookies;
  } catch {}
  try {
    const env = fs.readFileSync(ENV_PATH, "utf-8");
    const m = env.match(/^MIMO_CONSOLE_COOKIES=(.+)$/m);
    if (m) return m[1].trim();
  } catch {}
  return "";
}

function saveCookies(cookies) {
  let store = {};
  try { store = JSON.parse(fs.readFileSync(STORE_PATH, "utf-8")); } catch {}
  store.cookies = cookies;
  fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2));
}

// --- History persistence ---
function loadHistory() {
  try {
    const store = JSON.parse(fs.readFileSync(STORE_PATH, "utf-8"));
    return store.history || [];
  } catch { return []; }
}

function saveHistory(history) {
  let store = {};
  try { store = JSON.parse(fs.readFileSync(STORE_PATH, "utf-8")); } catch {}
  // Keep 90 days
  const cutoff = Date.now() - 90 * 24 * 60 * 60 * 1000;
  store.history = history.filter((h) => h.ts > cutoff);
  fs.writeFileSync(STORE_PATH, JSON.stringify(store, null, 2));
}

// --- MiMo API fetch ---
function mimoFetch(apiPath) {
  return new Promise((resolve, reject) => {
    const cookies = loadCookies();
    const reqUrl = new URL(apiPath, MIMO_BASE);
    const options = {
      hostname: reqUrl.hostname,
      path: reqUrl.pathname + reqUrl.search,
      method: "GET",
      headers: {
        Cookie: cookies,
        "User-Agent": "TokenMonitor/1.0",
        Accept: "application/json",
      },
    };
    const req = https.request(options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try {
          resolve({ status: res.statusCode, body: JSON.parse(data) });
        } catch {
          resolve({ status: res.statusCode, body: { raw: data.slice(0, 500) } });
        }
      });
    });
    req.on("error", reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error("timeout")); });
    req.end();
  });
}

// --- Poll and store ---
let latestData = null;

async function pollUsage() {
  try {
    const [detail, usage] = await Promise.all([
      mimoFetch("/api/v1/tokenPlan/detail"),
      mimoFetch("/api/v1/tokenPlan/usage"),
    ]);

    if (detail.status === 401 || usage.status === 401) {
      console.error("[poll] Cookie expired or invalid (401)");
      latestData = { error: "cookie_expired", ts: Date.now() };
      return;
    }

    const now = Date.now();
    const record = {
      ts: now,
      detail: detail.body,
      usage: usage.body,
    };

    latestData = record;

    const history = loadHistory();
    history.push({ ts: now, detail: detail.body, usage: usage.body });
    saveHistory(history);

    // Extract summary
    const d = detail.body;
    console.log(`[poll] ${new Date(now).toLocaleString()} - fetched OK`);
    if (d && d.data) {
      console.log(`  plan: ${d.data.planName || "?"}, used: ${d.data.usedCredits || "?"}/${d.data.totalCredits || "?"}`);
    }
  } catch (err) {
    console.error("[poll] error:", err.message);
  }
}

// --- Static file serving ---
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css",
  ".js": "application/javascript",
  ".json": "application/json",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

function serveStatic(req, res, filePath) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404); res.end("Not found"); return;
    }
    const ext = path.extname(filePath);
    res.writeHead(200, { "Content-Type": MIME[ext] || "text/plain" });
    res.end(data);
  });
}

// --- HTTP server ---
const server = http.createServer(async (req, res) => {
  const fullUrl = new URL(req.url, `http://localhost:${PORT}`);
  const pathname = fullUrl.pathname;
  const query = Object.fromEntries(fullUrl.searchParams);

  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }

  // API: get latest data
  if (pathname === "/api/usage") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify(latestData));
    return;
  }

  // API: get history
  if (pathname === "/api/history") {
    const history = loadHistory();
    const days = Number(query.days || 7);
    const cutoff = Date.now() - days * 24 * 60 * 60 * 1000;
    const filtered = history.filter((h) => h.ts > cutoff);
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify(filtered));
    return;
  }

  // API: save cookies
  if (pathname === "/api/cookies" && req.method === "POST") {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      try {
        const { cookies } = JSON.parse(body);
        saveCookies(cookies);
        pollUsage(); // immediate re-poll
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ ok: true }));
      } catch {
        res.writeHead(400); res.end("bad request");
      }
    });
    return;
  }

  // API: force refresh
  if (pathname === "/api/refresh") {
    await pollUsage();
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify(latestData));
    return;
  }

  // Static files
  let filePath = pathname === "/" ? "/monitor.html" : pathname;
  filePath = path.join(__dirname, filePath);
  // Security: prevent path traversal
  if (!filePath.startsWith(__dirname)) { res.writeHead(403); res.end(); return; }
  serveStatic(req, res, filePath);
});

// --- Start ---
console.log(`Token Monitor starting on http://localhost:${PORT}`);
console.log(`Cookie source: ${loadCookies() ? ".env / store.json" : "(none - set MIMO_CONSOLE_COOKIES or use web UI)"}`);

pollUsage(); // initial poll
setInterval(pollUsage, POLL_INTERVAL);

server.listen(PORT, () => {
  console.log(`Dashboard: http://localhost:${PORT}`);
});
