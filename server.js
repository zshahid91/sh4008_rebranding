/**
 * server.js — deck-tools sidecar service
 *
 * Runs LibreOffice + poppler-utils + pptxgenjs in its OWN container,
 * separate from the n8n container. n8n's official image went distroless
 * in v2.x (no apk/apt-get, no shell), so LibreOffice can no longer be
 * installed directly into the n8n image. This sidecar avoids that
 * problem entirely: n8n just calls it over HTTP, regardless of which
 * n8n image/version is running.
 *
 * Since this service is reachable over the public internet (not a
 * private docker network), every request must include a matching
 * shared-secret header: x-deck-tools-secret. Set it via the
 * DECK_TOOLS_SECRET environment variable when running the container.
 *
 * Endpoints:
 *   GET  /health                (no auth required)
 *   POST /pptx-to-images        raw binary body = a .pptx (n8n HTTP Request
 *                                node, Body Content Type "binaryData")
 *                                -> { count, images: [{ filename, base64 }] }
 *   POST /pptx-extract-assets   raw binary body = a .pptx
 *                                -> { count, assets: [{ filename, base64 }] }
 *                                (embedded images from ppt/media/, e.g. the
 *                                logo and any pictures used by Type D/J slides)
 *   POST /build-deck            JSON { designSpec, slideContent, assets }
 *                                -> raw .pptx binary
 */

const express = require("express");
const { execFile } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const AdmZip = require("adm-zip");

const app = express();
app.use(express.json({ limit: "200mb" }));

const PORT = process.env.PORT || 4000;
const SHARED_SECRET = process.env.DECK_TOOLS_SECRET || "";

app.use((req, res, next) => {
  if (req.path === "/health") return next();
  if (!SHARED_SECRET) return next(); // no secret configured = auth disabled (local/dev use only)
  const provided = req.get("x-deck-tools-secret");
  if (provided !== SHARED_SECRET) {
    return res.status(401).json({ error: "Missing or incorrect x-deck-tools-secret header" });
  }
  next();
});

function run(cmd, args, cwd) {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, { cwd, maxBuffer: 1024 * 1024 * 64 }, (err, stdout, stderr) => {
      if (err) {
        err.stdout = stdout;
        err.stderr = stderr;
        return reject(err);
      }
      resolve({ stdout, stderr });
    });
  });
}

function tmpDir(prefix) {
  const dir = path.join(os.tmpdir(), `${prefix}-${crypto.randomUUID()}`);
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

app.get("/health", (req, res) => res.json({ ok: true }));

// ----------------------------------------------------------------------
// POST /pptx-to-images
// ----------------------------------------------------------------------
app.post("/pptx-to-images", express.raw({ type: "*/*", limit: "200mb" }), async (req, res) => {
  if (!req.body || !req.body.length) return res.status(400).json({ error: "Missing raw request body (the .pptx bytes)" });

  const dir = tmpDir("p2i");
  const pptxPath = path.join(dir, "input.pptx");
  fs.writeFileSync(pptxPath, req.body);

  try {
    await run("soffice", ["--headless", "--convert-to", "pdf", "--outdir", dir, pptxPath]);
    const pdfPath = path.join(dir, "input.pdf");
    if (!fs.existsSync(pdfPath)) throw new Error("LibreOffice did not produce a PDF");

    await run("pdftoppm", ["-jpeg", "-r", "150", pdfPath, path.join(dir, "slide")]);

    const files = fs
      .readdirSync(dir)
      .filter((f) => f.startsWith("slide") && f.endsWith(".jpg"))
      .sort();

    const images = files.map((f) => ({
      filename: f,
      base64: fs.readFileSync(path.join(dir, f)).toString("base64"),
    }));

    res.json({ count: images.length, images });
  } catch (err) {
    res.status(500).json({ error: String(err.message || err), stderr: err.stderr ? String(err.stderr) : undefined });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

// ----------------------------------------------------------------------
// POST /pptx-extract-assets
// ----------------------------------------------------------------------
app.post("/pptx-extract-assets", express.raw({ type: "*/*", limit: "200mb" }), async (req, res) => {
  if (!req.body || !req.body.length) return res.status(400).json({ error: "Missing raw request body (the .pptx bytes)" });

  try {
    const zip = new AdmZip(req.body);
    const entries = zip.getEntries().filter((e) => e.entryName.startsWith("ppt/media/") && !e.isDirectory);
    const assets = entries.map((e) => ({
      filename: path.basename(e.entryName),
      base64: e.getData().toString("base64"),
    }));
    res.json({ count: assets.length, assets });
  } catch (err) {
    res.status(500).json({ error: String(err.message || err) });
  }
});

// ----------------------------------------------------------------------
// POST /build-deck
// ----------------------------------------------------------------------
app.post("/build-deck", async (req, res) => {
  const { designSpec, slideContent, assets } = req.body || {};
  if (!designSpec || !slideContent) {
    return res.status(400).json({ error: "Body must include designSpec and slideContent" });
  }

  const dir = tmpDir("build");
  const assetsDir = path.join(dir, "assets");
  fs.mkdirSync(assetsDir, { recursive: true });

  try {
    (assets || []).forEach((a) => {
      fs.writeFileSync(path.join(assetsDir, a.filename), Buffer.from(a.base64, "base64"));
    });

    const specPath = path.join(dir, "design_spec.json");
    const contentPath = path.join(dir, "slide_content.json");
    const outPath = path.join(dir, "output.pptx");
    fs.writeFileSync(specPath, JSON.stringify(designSpec));
    fs.writeFileSync(contentPath, JSON.stringify(slideContent));

    await run("node", [path.join(__dirname, "build_deck.js"), specPath, contentPath, outPath, assetsDir]);

    if (!fs.existsSync(outPath)) throw new Error("build_deck.js did not produce an output file");

    const buf = fs.readFileSync(outPath);
    res.set("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.set("Content-Disposition", 'attachment; filename="restyled.pptx"');
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: String(err.message || err), stderr: err.stderr ? String(err.stderr) : undefined });
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

app.listen(PORT, () => console.log(`deck-tools service listening on :${PORT}`));
