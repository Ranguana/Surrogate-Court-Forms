const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const https = require("https");
const { spawn, exec } = require("child_process");

const FLASK_PORT = 52845;
const GITHUB_REPO = "Ranguana/Surrogate-Court-Forms";
let mainWindow;
let flaskProcess;
let flaskFailed = false;
let flaskError = "";

function getBundledAppDir() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, "app");
  }
  return __dirname;
}

function getLiveAppDir() {
  return path.join(app.getPath("userData"), "live_app");
}

function getAppDir() {
  const live = getLiveAppDir();
  // Use live (updatable) copy if it exists, otherwise fall back to bundled
  if (fs.existsSync(path.join(live, "app.py"))) {
    return live;
  }
  return getBundledAppDir();
}

function sendStatus(msg) {
  console.log("[STATUS]", msg);
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.executeJavaScript(
      `window.__statusUpdate && window.__statusUpdate(${JSON.stringify(msg)})`
    );
  }
}

function notifyFlaskFailed(msg) {
  flaskFailed = true;
  console.error("[FAILED]", msg);
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.executeJavaScript(
      `window.__flaskFailed && window.__flaskFailed(${JSON.stringify(msg)})`
    );
  }
}

function run(cmd, timeout = 120000) {
  return new Promise((resolve, reject) => {
    exec(cmd, { timeout }, (err, stdout, stderr) => {
      if (err) {
        console.error("[CMD FAIL]", cmd, "\n", stderr);
        reject(new Error(stderr || err.message));
      } else {
        resolve(stdout.trim());
      }
    });
  });
}

async function findPython() {
  const candidates = [
    "/usr/bin/python3",
    "/usr/local/bin/python3",
    "/opt/homebrew/bin/python3",
    "/Library/Frameworks/Python.framework/Versions/Current/bin/python3",
  ];
  for (const p of candidates) {
    try {
      if (fs.existsSync(p)) {
        await run(`"${p}" --version`, 5000);
        console.log("[PYTHON] Found:", p);
        return p;
      }
    } catch (e) { /* skip */ }
  }
  return null;
}

function getVenvDir() {
  return path.join(app.getPath("userData"), "python_env");
}

async function setupPythonEnv(systemPython) {
  const venvDir = getVenvDir();
  const venvPython = path.join(venvDir, "bin", "python3");
  const marker = path.join(venvDir, ".deps_installed");

  // If venv and deps already exist, return immediately
  if (fs.existsSync(venvPython) && fs.existsSync(marker)) {
    console.log("[PYTHON] Existing venv found");
    return venvPython;
  }

  // Create venv
  if (!fs.existsSync(venvPython)) {
    sendStatus("Setting up Python environment (first time only)...");
    await run(`"${systemPython}" -m venv "${venvDir}"`, 30000);
  }

  // Install dependencies
  sendStatus("Installing dependencies — this may take a minute...");
  const deps = "flask python-docx pypdf pdfplumber pymupdf openpyxl python-dotenv anthropic pytesseract pdf2image";
  await run(`"${venvPython}" -m pip install --upgrade pip --quiet`, 60000);
  await run(`"${venvPython}" -m pip install ${deps} --quiet`, 180000);

  // Install Tesseract & Poppler system binaries for OCR (macOS via Homebrew)
  try {
    await run("which tesseract", 5000);
    console.log("[OCR] Tesseract already installed");
  } catch (e) {
    try {
      await run("which brew", 5000);
      sendStatus("Installing OCR support (tesseract)...");
      await run("brew install tesseract --quiet", 120000);
      console.log("[OCR] Tesseract installed via Homebrew");
    } catch (brewErr) {
      console.log("[OCR] Homebrew not available — Tesseract OCR skipped, Claude vision will be used as fallback");
    }
  }
  try {
    await run("which pdftoppm", 5000);
    console.log("[OCR] Poppler already installed");
  } catch (e) {
    try {
      await run("which brew", 5000);
      sendStatus("Installing OCR support (poppler)...");
      await run("brew install poppler --quiet", 120000);
      console.log("[OCR] Poppler installed via Homebrew");
    } catch (brewErr) {
      console.log("[OCR] Homebrew not available — Poppler skipped, Claude vision will be used as fallback");
    }
  }

  fs.writeFileSync(marker, new Date().toISOString());
  console.log("[PYTHON] Dependencies installed");
  return venvPython;
}

function launchFlask(python, appDir) {
  sendStatus("Starting server...");
  flaskProcess = spawn(python, ["app.py"], {
    cwd: appDir,
    env: { ...process.env, PYTHONUNBUFFERED: "1" },
  });
  flaskProcess.stdout.on("data", (d) => process.stdout.write(d));
  flaskProcess.stderr.on("data", (d) => {
    process.stderr.write(d);
    flaskError += d.toString();
  });
  flaskProcess.on("exit", (code) => {
    if (code !== 0 && code !== null) {
      notifyFlaskFailed(flaskError || `Server exited with code ${code}`);
    }
  });
  flaskProcess.on("error", (err) => {
    flaskError = err.message;
    notifyFlaskFailed("Could not start server: " + err.message);
  });
}

// ── Auto-update: pull latest source from GitHub ─────────────────────────────

function httpGet(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, { headers: { "User-Agent": "ProbateAssistant" } }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return httpGet(res.headers.location).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      const chunks = [];
      res.on("data", (c) => chunks.push(c));
      res.on("end", () => resolve(Buffer.concat(chunks)));
    });
    req.on("error", reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error("timeout")); });
  });
}

function getLocalVersion() {
  // Read version from the live app dir, then bundled, then fallback
  for (const dir of [getLiveAppDir(), getBundledAppDir()]) {
    const appPy = path.join(dir, "app.py");
    if (fs.existsSync(appPy)) {
      const content = fs.readFileSync(appPy, "utf8");
      const m = content.match(/APP_VERSION\s*=\s*"([^"]+)"/);
      if (m) return m[1];
    }
  }
  return "0.0.0";
}

async function checkAndUpdate() {
  try {
    sendStatus("Checking for updates...");
    const releaseData = await httpGet(`https://api.github.com/repos/${GITHUB_REPO}/releases/latest`);
    const release = JSON.parse(releaseData.toString());
    const latest = (release.tag_name || "").replace(/^v/, "");
    const current = getLocalVersion();

    if (!latest || latest === current) {
      console.log(`[UPDATE] Up to date (v${current})`);
      return;
    }

    console.log(`[UPDATE] New version available: v${latest} (current: v${current})`);
    sendStatus(`Updating to v${latest}...`);

    // Download the source zip from the release tag
    const zipUrl = `https://github.com/${GITHUB_REPO}/archive/refs/tags/${release.tag_name}.zip`;
    const zipData = await httpGet(zipUrl);

    // Write zip to temp file and extract
    const tmpZip = path.join(app.getPath("temp"), "probate_update.zip");
    fs.writeFileSync(tmpZip, zipData);

    const liveDir = getLiveAppDir();
    const tmpExtract = path.join(app.getPath("temp"), "probate_extract");

    // Clean previous extract
    if (fs.existsSync(tmpExtract)) {
      fs.rmSync(tmpExtract, { recursive: true, force: true });
    }

    await run(`unzip -o -q "${tmpZip}" -d "${tmpExtract}"`, 30000);

    // Find the extracted folder (GitHub zips have a top-level folder like Repo-Name-tag/)
    const extracted = fs.readdirSync(tmpExtract);
    const srcDir = path.join(tmpExtract, extracted[0]);

    // Copy updated files to live app dir
    if (fs.existsSync(liveDir)) {
      fs.rmSync(liveDir, { recursive: true, force: true });
    }
    fs.mkdirSync(liveDir, { recursive: true });

    // Copy all needed files/dirs
    const items = ["app.py", "generators.py", "requirements.txt", "static", "templates",
                   "Accounting", "Probate-_NY_Court_Forms.pdf", "admin_ancil.pdf",
                   "Petition_for_Non-Domciliary_Letters_of_Admin.pdf", "login.html",
                   "preload.js", "favicon.svg"];
    for (const item of items) {
      const src = path.join(srcDir, item);
      const dst = path.join(liveDir, item);
      if (fs.existsSync(src)) {
        await run(`cp -R "${src}" "${dst}"`, 10000);
      }
    }

    // Copy .env from bundled app (has API key, not in git)
    const bundledEnv = path.join(getBundledAppDir(), ".env");
    const liveEnv = path.join(liveDir, ".env");
    if (fs.existsSync(bundledEnv)) {
      fs.copyFileSync(bundledEnv, liveEnv);
    }

    // Copy contacts/cases from bundled dir if not in live dir yet
    for (const f of ["cases.json", "contacts.json"]) {
      const dst = path.join(liveDir, f);
      if (!fs.existsSync(dst)) {
        const bundled = path.join(getBundledAppDir(), f);
        if (fs.existsSync(bundled)) {
          fs.copyFileSync(bundled, dst);
        } else {
          fs.writeFileSync(dst, "{}");
        }
      }
    }

    // Clean up
    try { fs.unlinkSync(tmpZip); } catch (e) { /* ok */ }
    try { fs.rmSync(tmpExtract, { recursive: true, force: true }); } catch (e) { /* ok */ }

    // Force re-install deps in case requirements changed
    const marker = path.join(getVenvDir(), ".deps_installed");
    try { fs.unlinkSync(marker); } catch (e) { /* ok */ }

    console.log(`[UPDATE] Updated to v${latest}`);
    sendStatus(`Updated to v${latest}! Starting...`);
  } catch (e) {
    console.log(`[UPDATE] Check failed (continuing with current version): ${e.message}`);
  }
}

async function startFlask() {
  // Check for updates before anything else
  await checkAndUpdate();

  const appDir = getAppDir();

  sendStatus("Finding Python...");
  const systemPython = await findPython();
  if (!systemPython) {
    notifyFlaskFailed(
      "Python 3 not found. Please install from python.org/downloads and relaunch."
    );
    return;
  }

  let python;
  try {
    python = await setupPythonEnv(systemPython);
  } catch (e) {
    notifyFlaskFailed("Setup failed: " + e.message);
    return;
  }

  launchFlask(python, appDir);
}

ipcMain.handle("flask-status", () => {
  return { failed: flaskFailed, error: flaskError };
});

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: "NY Surrogate's Court Probate Assistant",
    webPreferences: {
      preload: path.join(getAppDir(), "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  mainWindow.loadFile(path.join(getAppDir(), "login.html"));
}

app.whenReady().then(() => {
  createWindow();
  mainWindow.webContents.on("did-finish-load", () => {
    startFlask();
  });
});

app.on("window-all-closed", () => {
  if (flaskProcess) flaskProcess.kill();
  app.quit();
});

app.on("before-quit", () => {
  if (flaskProcess) flaskProcess.kill();
});
