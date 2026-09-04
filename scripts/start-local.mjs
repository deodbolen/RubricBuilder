import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const url = "http://127.0.0.1:4173/";
const nodeMajor = Number.parseInt(process.versions.node.split(".")[0], 10);

function log(message = "") {
  console.log(message);
}

function run(command, args, options = {}) {
  return spawnSync(command, args, {
    cwd: root,
    stdio: "inherit",
    shell: process.platform === "win32",
    ...options,
  });
}

async function hasDependencies() {
  try {
    await import("exceljs");
    return true;
  } catch {
    return false;
  }
}

async function waitForServer(timeoutMs = 12000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return true;
    } catch {
      // The server is still starting.
    }
    await new Promise((resolveWait) => setTimeout(resolveWait, 350));
  }
  return false;
}

function openBrowser() {
  if (process.platform === "darwin") {
    spawn("open", [url], { detached: true, stdio: "ignore" }).unref();
    return;
  }
  if (process.platform === "win32") {
    spawn("cmd", ["/c", "start", "", url], { detached: true, stdio: "ignore" }).unref();
    return;
  }
  spawn("xdg-open", [url], { detached: true, stdio: "ignore" }).unref();
}

function printNodeInstallHelp() {
  log("Node.js 20 or newer is required to run Rubric Builder.");
  log("Install the LTS version from https://nodejs.org/, then run this starter again.");
}

if (nodeMajor < 20 || Number.isNaN(nodeMajor)) {
  printNodeInstallHelp();
  process.exit(1);
}

log("Starting Rubric Builder...");

if (!(await hasDependencies())) {
  log("Installing first-run dependencies. This can take a minute...");
  const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";
  const install = run(npmCommand, ["install"]);
  if (install.status !== 0 || !(await hasDependencies())) {
    log("");
    log("Dependency install did not finish successfully.");
    log("Check your internet connection, then try this starter again.");
    process.exit(1);
  }
}

if (!existsSync(resolve(root, "src", "Demo Rubric Template.xlsx"))) {
  log("Missing src/Demo Rubric Template.xlsx. Make sure the full project folder was downloaded.");
  process.exit(1);
}

if (await waitForServer(1000)) {
  log(`Rubric Builder is already running. Opening ${url}`);
  openBrowser();
  process.exit(0);
}

const server = spawn(process.execPath, ["server.mjs"], {
  cwd: root,
  stdio: "inherit",
});

server.on("exit", (code) => {
  process.exit(code ?? 0);
});

if (await waitForServer()) {
  log(`Opening ${url}`);
  openBrowser();
} else {
  log(`The server is taking longer than expected. Try opening ${url} manually.`);
}
