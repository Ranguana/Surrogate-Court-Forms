const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const https = require("https");
const { spawn, exec } = require("child_process");
const crypto = require("crypto");

const FLASK_PORT = 52845;
const GITHUB_REPO = "Ranguana/Surrogate-Court-Forms";
let mainWindow;
let flaskProcess;
let flaskFailed = false;
let flaskError = "";

// ── Path helpers ────────────────────────────────────────────────────────────

function getServerBinary() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, "probate-server");
  }
  // Dev mode: use local dist/ output from PyInstaller
  const local = path.join(__dirname, "dist", "probate-server");
  if (fs.existsSync(local)) return local;
  return null;
}

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
  if (fs.existsSync(path.join(live, "app.py"))) {
    return live;
  }
  return getBundledAppDir();
}

// ── Status messaging ────────────────────────────────────────────────────────

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
  flaskError = msg;
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

// ── Server launcher ─────────────────────────────────────────────────────────

function launchServer(binaryPath, appDir) {
  sendStatus("Starting server...");

  // The frozen binary takes the app directory as its argument
  // It loads app.py from that directory at runtime
  flaskProcess = spawn(binaryPath, [appDir], {
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
    notifyFlaskFailed("Could not start server: " + err.message);
  });
}

// ── Auto-update (source files only — not the binary) ────────────────────────

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
  const maxRetries = 3;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      sendStatus("Checking for updates...");
      const releaseData = await httpGet(`https://api.github.com/repos/${GITHUB_REPO}/releases/latest`);
      const release = JSON.parse(releaseData.toString());
      const latest = (release.tag_name || "").replace(/^v/, "");
      const current = getLocalVersion();

      if (!latest || latest === current) {
        console.log(`[UPDATE] Up to date (v${current})`);
        sendStatus(`v${current} — up to date`);
        return;
      }

      console.log(`[UPDATE] New version available: v${latest} (current: v${current})`);
      sendStatus(`Updating to v${latest}...`);

      const zipUrl = `https://github.com/${GITHUB_REPO}/archive/refs/tags/${release.tag_name}.zip`;
      const zipData = await httpGet(zipUrl);

      const tmpZip = path.join(app.getPath("temp"), "probate_update.zip");
      fs.writeFileSync(tmpZip, zipData);

      const liveDir = getLiveAppDir();
      const tmpExtract = path.join(app.getPath("temp"), "probate_extract");

      if (fs.existsSync(tmpExtract)) {
        fs.rmSync(tmpExtract, { recursive: true, force: true });
      }

      await run(`unzip -o -q "${tmpZip}" -d "${tmpExtract}"`, 30000);

      const extracted = fs.readdirSync(tmpExtract);
      const srcDir = path.join(tmpExtract, extracted[0]);

      if (fs.existsSync(liveDir)) {
        fs.rmSync(liveDir, { recursive: true, force: true });
      }
      fs.mkdirSync(liveDir, { recursive: true });

      const items = [
        "app.py", "generators.py", "requirements.txt",
        "static", "templates", "Accounting",
        "Probate-_NY_Court_Forms.pdf", "admin_ancil.pdf",
        "Petition_for_Non-Domciliary_Letters_of_Admin.pdf",
        "login.html", "preload.js", "favicon.svg",
        "field_mappings.py",
      ];
      for (const item of items) {
        const src = path.join(srcDir, item);
        const dst = path.join(liveDir, item);
        if (fs.existsSync(src)) {
          await run(`cp -R "${src}" "${dst}"`, 10000);
        }
      }

      const bundledEnv = path.join(getBundledAppDir(), ".env");
      const liveEnv = path.join(liveDir, ".env");
      if (fs.existsSync(bundledEnv)) {
        fs.copyFileSync(bundledEnv, liveEnv);
      }

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

      try { fs.unlinkSync(tmpZip); } catch (e) { /* ok */ }
      try { fs.rmSync(tmpExtract, { recursive: true, force: true }); } catch (e) { /* ok */ }

      console.log(`[UPDATE] Updated to v${latest}`);
      sendStatus(`Updated to v${latest}! Starting...`);
      return;
    } catch (e) {
      console.log(`[UPDATE] Attempt ${attempt}/${maxRetries} failed: ${e.message}`);
      if (attempt < maxRetries) {
        sendStatus(`Update check failed, retrying (${attempt}/${maxRetries})...`);
        await new Promise(r => setTimeout(r, 2000));
      } else {
        const current = getLocalVersion();
        sendStatus(`⚠ Could not check for updates — running v${current}`);
        console.log(`[UPDATE] All retries failed: ${e.message}`);
      }
    }
  }
}

// ── Main startup ────────────────────────────────────────────────────────────

async function startFlask() {
  await checkAndUpdate();

  const binary = getServerBinary();
  const appDir = getAppDir();

  if (!binary || !fs.existsSync(binary)) {
    notifyFlaskFailed(
      "The application appears to be damaged. Please re-download and reinstall Probate HQ."
    );
    return;
  }

  if (!fs.existsSync(path.join(appDir, "app.py"))) {
    notifyFlaskFailed(
      "Application files are missing. Please re-download and reinstall Probate HQ."
    );
    return;
  }

  console.log("[SERVER] Binary:", binary);
  console.log("[SERVER] App dir:", appDir);
  launchServer(binary, appDir);
}

// ── IPC handlers ────────────────────────────────────────────────────────────

ipcMain.handle("flask-status", () => {
  return { failed: flaskFailed, error: flaskError };
});

function getPasswordFile() {
  return path.join(app.getPath("userData"), "password.hash");
}

function sha256(str) {
  return crypto.createHash("sha256").update(str).digest("hex");
}

ipcMain.handle("has-password", () => {
  return fs.existsSync(getPasswordFile());
});

ipcMain.handle("login", async (_event, password) => {
  const file = getPasswordFile();
  if (!fs.existsSync(file)) return { ok: false, error: "No password set" };
  const saved = fs.readFileSync(file, "utf8").trim();
  if (sha256(password) === saved) return { ok: true };
  return { ok: false, error: "Incorrect password" };
});

ipcMain.handle("setup-password", async (_event, password) => {
  if (password.length < 4) return { ok: false, error: "Password must be at least 4 characters" };
  fs.writeFileSync(getPasswordFile(), sha256(password), "utf8");
  return { ok: true };
});

// ── Window ──────────────────────────────────────────────────────────────────

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: "NY Surrogate's Court — Probate HQ",
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