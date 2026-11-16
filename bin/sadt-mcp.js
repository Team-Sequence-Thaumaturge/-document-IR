#!/usr/bin/env node
"use strict";

const { spawn, spawnSync } = require("child_process");
const path = require("path");

function findPython() {
  const candidates = [];
  if (process.env.SADT_PYTHON) {
    candidates.push([process.env.SADT_PYTHON, []]);
  }
  candidates.push(["python3", []], ["python", []]);
  if (process.platform === "win32") {
    candidates.push(["py", ["-3"]]);
  }

  for (const [cmd, baseArgs] of candidates) {
    const check = spawnSync(
      cmd,
      [...baseArgs, "-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"],
      { stdio: "ignore", windowsHide: true }
    );
    if (!check.error && check.status === 0) {
      return { cmd, baseArgs };
    }
  }
  return null;
}

const py = findPython();
if (!py) {
  process.stderr.write(
    "[sadt-mcp] Error: Python 3.10+ is required on PATH (or set SADT_PYTHON).\n"
  );
  process.exit(1);
}

const serverScript = path.join(__dirname, "..", "mcp_server.py");
const child = spawn(
  py.cmd,
  [...py.baseArgs, "-u", serverScript, ...process.argv.slice(2)],
  {
    stdio: "inherit",
    windowsHide: true,
    env: {
      ...process.env,
      PYTHONUNBUFFERED: "1",
      PYTHONIOENCODING: "utf-8"
    }
  }
);

for (const sig of ["SIGINT", "SIGTERM"]) {
  process.on(sig, () => {
    if (!child.killed) {
      child.kill(sig);
    }
  });
}

child.on("error", (err) => {
  process.stderr.write(`[sadt-mcp] Failed to start Python server: ${err.message}\n`);
  process.exit(1);
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.exit(1);
  }
  process.exit(code == null ? 0 : code);
});
