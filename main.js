const { app, BrowserWindow } = require("electron");
const path = require("path");
const { spawn } = require("child_process");

const FLASK_PORT = 8080;
let mainWindow;
let flaskProcess;

function getAppDir() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, "app");
  }
  return __dirname;
}

function startFlask() {
  const appDir = getAppDir();
  let bundleDir;
  if (app.isPackaged) {
    bundleDir = path.join(process.resourcesPath, "python_bundle");
  } else {
    bundleDir = path.join(__dirname, "python_bundle");
  }
  const python = path.join(bundleDir, "bin", "python3");
  const pythonHome = bundleDir;
  const pythonPath = path.join(bundleDir, "lib", "python3.13");
  flaskProcess = spawn(python, ["app.py"], {
    cwd: appDir,
    env: {
      ...process.env,
      PYTHONUNBUFFERED: "1",
      PYTHONHOME: pythonHome,
      PYTHONPATH: pythonPath,
    },
  });
  flaskProcess.stdout.on("data", (d) => process.stdout.write(d));
  flaskProcess.stderr.on("data", (d) => process.stderr.write(d));
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    title: "NY Surrogate's Court Probate Assistant",
  });
  mainWindow.loadFile(path.join(getAppDir(), "login.html"));
}

app.whenReady().then(() => {
  startFlask();
  createWindow();
});

app.on("window-all-closed", () => {
  if (flaskProcess) flaskProcess.kill();
  app.quit();
});

app.on("before-quit", () => {
  if (flaskProcess) flaskProcess.kill();
});
