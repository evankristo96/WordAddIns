// Tiny HTTPS static server for sideloading the add-in on https://localhost:3000
// Word requires add-in pages to be served over HTTPS.
//
// One-time setup (installs a trusted local certificate):
//   npx office-addin-dev-certs install
// Then run:
//   node serve.js

const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = 3000;
const ROOT = __dirname;
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");
const certPath = path.join(CERT_DIR, "localhost.crt");
const keyPath = path.join(CERT_DIR, "localhost.key");

if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
  console.error("Certificate not found in " + CERT_DIR);
  console.error("Run this once, then try again:\n  npx office-addin-dev-certs install");
  process.exit(1);
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
};

const server = https.createServer(
  { cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) },
  (req, res) => {
    const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
    const rel = urlPath === "/" ? "/taskpane.html" : urlPath;
    const filePath = path.normalize(path.join(ROOT, rel));

    // stay inside the project folder
    if (!filePath.startsWith(ROOT + path.sep)) {
      res.writeHead(403);
      return res.end("Forbidden");
    }

    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        return res.end("Not found");
      }
      res.writeHead(200, {
        "Content-Type": MIME[path.extname(filePath)] || "application/octet-stream",
        "Cache-Control": "no-store",
      });
      res.end(data);
    });
  }
);

server.listen(PORT, () => {
  console.log("Quick Text add-in running at https://localhost:" + PORT);
  console.log("Leave this window open while you use the add-in. Press Ctrl+C to stop.");
});
