#!/usr/bin/env node
/**
 * Bumps the patch version in package.json + app.py, then builds the DMG.
 *
 * Usage:
 *   npm run build            → auto-bumps patch (1.3.0 → 1.3.1) and builds
 *   npm run build -- 1.4.0   → sets exact version and builds
 *   npm run build:nobump     → builds without changing version
 */
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const pkgPath = path.join(__dirname, "package.json");
const appPyPath = path.join(__dirname, "app.py");

// Read current version
const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf8"));
const currentVersion = pkg.version;

// Determine new version
let newVersion;
const arg = process.argv[2];
if (arg && /^\d+\.\d+\.\d+$/.test(arg)) {
  newVersion = arg;
} else {
  // Auto-bump patch
  const parts = currentVersion.split(".").map(Number);
  parts[2] += 1;
  newVersion = parts.join(".");
}

console.log(`\n  Version: ${currentVersion} → ${newVersion}\n`);

// Update package.json
pkg.version = newVersion;
fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + "\n");

// Update app.py
let appPy = fs.readFileSync(appPyPath, "utf8");
appPy = appPy.replace(/APP_VERSION\s*=\s*"[^"]+"/, `APP_VERSION = "${newVersion}"`);
fs.writeFileSync(appPyPath, appPy);

// Build
console.log("  Building DMG...\n");
try {
  execSync("npx electron-builder --mac", { stdio: "inherit", cwd: __dirname });
  const dmg = path.join(__dirname, "dist", `Probate HQ-${newVersion}-arm64.dmg`);
  console.log(`\n  ✓ DMG ready: dist/Probate HQ-${newVersion}-arm64.dmg\n`);
} catch (e) {
  process.exit(1);
}
