const fs = require("fs");
const path = require("path");
const os = require("os");
const https = require("https");
const express = require("express");

const PORT = process.env.PORT || 3000;

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

// Stub API endpoints
app.post("/api/improve", (req, res) => {
  const { selectedText } = req.body || {};
  const improvedText =
    typeof selectedText === "string" ? selectedText.toUpperCase() : "";
  res.json({ improvedText });
});

app.post("/api/review", (req, res) => {
  const { text } = req.body || {};
  const suggestions = [];
  if (typeof text === "string" && text.toLowerCase().includes("utilize")) {
    suggestions.push({
      anchor: "utilize",
      replacement: "use",
      reason: "Clarity",
    });
  }
  res.json({ suggestions });
});

const { key, cert } = ensureCertificates();

https
  .createServer({ key, cert }, app)
  .listen(PORT, () =>
    console.log(`HTTPS server running at https://localhost:${PORT}`)
  );
