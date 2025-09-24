const fs = require("fs");
const path = require("path");
const os = require("os");
const https = require("https");
const express = require("express");
const JSON5 = require("json5");

const PORT = process.env.PORT || 3001;
const AI_API_URL =
  process.env.AI_API_URL || "http://localhost:3000/api/office-addin";
const IMPROVE_URL =
  process.env.IMPROVE_URL || "http://localhost:3000/api/office-addin/improve"; // absolute or relative
const REVIEW_URL =
  process.env.REVIEW_URL || "http://localhost:3000/api/office-addin/review"; // absolute or relative
// const AI_API_KEY = process.env.AI_API_KEY || null; // disabled per request

function getCertificatePaths() {
  const homeDir = os.homedir();
  const certDir = path.join(homeDir, ".office-addin-dev-certs");
  return {
    certPath: path.join(certDir, "localhost.crt"),
    keyPath: path.join(certDir, "localhost.key"),
  };
}

function ensureCertificates() {
  const { certPath, keyPath } = getCertificatePaths();
  const certExists = fs.existsSync(certPath);
  const keyExists = fs.existsSync(keyPath);
  if (!certExists || !keyExists) {
    console.error(
      "Developer HTTPS certificates not found. Run 'npm run dev-cert' once, then restart."
    );
    process.exit(1);
  }
  return {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
}

const app = express();
app.use(express.json());

// Serve task pane static assets
app.use(express.static(path.join(__dirname, "public")));
app.get("/", (_req, res) => {
  res.redirect("/taskpane.html");
});

async function callAI(endpointOrPath, payload) {
  if (!AI_API_URL) {
    // Allow absolute URLs even if AI_API_URL is not set
    const isAbsolute = /^https?:\/\//i.test(endpointOrPath || "");
    if (!isAbsolute) {
      throw new Error(
        "AI_API_URL is not configured and endpoint is not absolute"
      );
    }
  }
  const url = /^https?:\/\//i.test(endpointOrPath)
    ? endpointOrPath
    : new URL(endpointOrPath, AI_API_URL).toString();
  const headers = { "Content-Type": "application/json" };
  // if (AI_API_KEY) headers["Authorization"] = `Bearer ${AI_API_KEY}`; // disabled per request
  const resp = await fetch(url, {
    method: "POST",
    headers,
    body: JSON.stringify(payload || {}),
  });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`AI API ${resp.status}: ${text}`);
  }
  const text = await resp.text();
  return text;
}

// Improve selection → proxy to AI or fallback to stub
app.post("/api/improve", async (req, res) => {
  const { selectedText } = req.body || {};
  if (true) {
    try {
      const target = IMPROVE_URL || "/improve";
      console.log(`[proxy] /api/improve → ${target}`);
      let raw = await callAI(target, { text: selectedText });
      let improvedText = "";
      try {
        const parsed = JSON.parse(raw);
        improvedText =
          typeof parsed === "string" ? parsed : parsed?.improvedText || "";
      } catch (_) {
        try {
          const parsed5 = JSON5.parse(raw);
          improvedText =
            typeof parsed5 === "string" ? parsed5 : parsed5?.improvedText || "";
        } catch (_) {
          improvedText = String(raw);
        }
      }
      return res.json({ improvedText });
    } catch (err) {
      console.error("[proxy] /api/improve failed, falling back to stub:", err);
      return res.status(502).json({ error: "Improve failed" });
    }
  }
  console.log("[stub] /api/improve using local uppercase fallback");
  const improvedText =
    typeof selectedText === "string" ? selectedText.toUpperCase() : "";
  return res.json({ improvedText });
});

// Review whole doc → proxy to AI or fallback to stub
app.post("/api/review", async (req, res) => {
  const { text } = req.body || {};
  if (REVIEW_URL || AI_API_URL) {
    try {
      const target = REVIEW_URL || "/review";
      console.log(`[proxy] /api/review → ${target}`);
      const raw = await callAI(target, { text });
      console.log("[api/review] raw →", raw);
      // Prefer suggestions list if present; otherwise treat as full document text
      try {
        const parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.suggestions)) {
          return res.json({ suggestions: parsed.suggestions });
        }
        if (typeof parsed === "string") {
          return res.json({ fullText: parsed });
        }
        if (parsed && typeof parsed.fullText === "string") {
          return res.json({ fullText: parsed.fullText });
        }
        if (parsed && typeof parsed.replacement === "string") {
          return res.json({ fullText: parsed.replacement });
        }
      } catch {}
      try {
        const parsed5 = JSON5.parse(raw);
        if (parsed5 && Array.isArray(parsed5.suggestions)) {
          return res.json({ suggestions: parsed5.suggestions });
        }
        if (typeof parsed5 === "string") {
          return res.json({ fullText: parsed5 });
        }
        if (parsed5 && typeof parsed5.fullText === "string") {
          return res.json({ fullText: parsed5.fullText });
        }
        if (parsed5 && typeof parsed5.replacement === "string") {
          return res.json({ fullText: parsed5.replacement });
        }
      } catch {}
      return res.json({ fullText: String(raw) });
    } catch (err) {
      console.error("[proxy] /api/review failed:", err);
      return res.status(502).json({ error: "Review failed" });
    }
  }
  // Stub fallback: return empty suggestions
  console.log("[stub] /api/review using local suggestion fallback");
  return res.json({ suggestions: [] });
});

const { key, cert } = ensureCertificates();

https
  .createServer({ key, cert }, app)
  .listen(PORT, () =>
    console.log(`HTTPS server running at https://localhost:${PORT}`)
  );
