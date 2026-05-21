const { app, BrowserWindow, ipcMain, shell } = require("electron");
const path = require("path");
const fs = require("fs");
const { spawn } = require("child_process");

const isDev = !app.isPackaged;
const baseDir = path.resolve(__dirname, "..", "..");
const runtimeDir = path.join(baseDir, "hma-desktop", ".runtime");

if (!fs.existsSync(runtimeDir)) {
  fs.mkdirSync(runtimeDir, { recursive: true });
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1320,
    height: 820,
    minWidth: 1180,
    minHeight: 720,
    title: "HMA Manager",
    backgroundColor: "#06122a",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  if (isDev) {
    win.loadURL("http://127.0.0.1:5173");
  } else {
    win.loadFile(path.join(__dirname, "..", "dist", "index.html"));
  }
}

function isInsideBase(targetPath) {
  const relative = path.relative(baseDir, targetPath);
  return relative && !relative.startsWith("..") && !path.isAbsolute(relative);
}

function limitOutput(text, max = 50000) {
  if (!text) return "";
  if (text.length <= max) return text;
  return text.slice(text.length - max);
}

function normalizeOutput(value) {
  return value ? value.toString("utf8") : "";
}

function walkFiles(dir, extensions, results = []) {
  if (!fs.existsSync(dir)) return results;

  const items = fs.readdirSync(dir, { withFileTypes: true });

  for (const item of items) {
    const full = path.join(dir, item.name);

    if (item.isDirectory()) {
      walkFiles(full, extensions, results);
    } else {
      const ext = path.extname(item.name).toLowerCase();
      if (extensions.includes(ext)) results.push(full);
    }
  }

  return results;
}

function readTextFileSafe(filePath, maxBytes = 250000) {
  try {
    const stat = fs.statSync(filePath);
    const fd = fs.openSync(filePath, "r");
    const bytesToRead = Math.min(stat.size, maxBytes);
    const start = Math.max(0, stat.size - bytesToRead);
    const buffer = Buffer.alloc(bytesToRead);

    fs.readSync(fd, buffer, 0, bytesToRead, start);
    fs.closeSync(fd);

    const prefix = stat.size > maxBytes ? "\n[ARCHIVO LARGO: mostrando solo el final]\n\n" : "";
    return prefix + buffer.toString("utf8");
  } catch (error) {
    return `No se pudo leer archivo: ${error.message}`;
  }
}

function collectGeneratedTextFiles(startMs) {
  const roots = [
    path.join(baseDir, "diagnosticos"),
    path.join(baseDir, "logs")
  ];

  const files = [];

  for (const root of roots) {
    const found = walkFiles(root, [".txt", ".log"]);

    for (const file of found) {
      try {
        const stat = fs.statSync(file);

        if (stat.mtimeMs >= startMs - 2500) {
          files.push({
            file,
            mtimeMs: stat.mtimeMs,
            size: stat.size
          });
        }
      } catch {}
    }
  }

  return files.sort((a, b) => b.mtimeMs - a.mtimeMs).slice(0, 4);
}

function appendGeneratedFilesToResult(result, startMs) {
  const generated = collectGeneratedTextFiles(startMs);

  if (!generated.length) return result;

  let extra = "\n\n============================================================\n";
  extra += "ARCHIVOS TXT/LOG GENERADOS DURANTE LA ACCION\n";
  extra += "============================================================\n";

  for (const item of generated) {
    extra += `\n\n--- ${item.file} ---\n`;
    extra += readTextFileSafe(item.file);
  }

  return {
    ...result,
    stdout: limitOutput((result.stdout || "") + extra)
  };
}

function runProcess(command, args = [], cwd = baseDir) {
  return new Promise((resolve) => {
    let stdout = "";
    let stderr = "";

    const child = spawn(command, args, {
      cwd,
      windowsHide: true,
      shell: false,
      env: {
        ...process.env,
        HMA_DASHBOARD_MODE: "1",
        HMA_NO_NOTEPAD: "1",
        HMA_NO_OPEN_TXT: "1"
      }
    });

    const timer = setTimeout(() => {
      try {
        child.kill("SIGTERM");
      } catch {}

      resolve({
        ok: false,
        code: -1,
        stdout: limitOutput(stdout),
        stderr: limitOutput(stderr + "\nTIMEOUT: el proceso supero 10 minutos.")
      });
    }, 1000 * 60 * 10);

    child.stdout.on("data", (data) => {
      stdout += normalizeOutput(data);
    });

    child.stderr.on("data", (data) => {
      stderr += normalizeOutput(data);
    });

    child.on("error", (error) => {
      clearTimeout(timer);

      resolve({
        ok: false,
        code: -1,
        stdout: limitOutput(stdout),
        stderr: limitOutput(stderr + "\n" + error.message)
      });
    });

    child.on("close", (code) => {
      clearTimeout(timer);

      resolve({
        ok: code === 0,
        code,
        stdout: limitOutput(stdout),
        stderr: limitOutput(stderr)
      });
    });
  });
}

function resolveSafeTarget(command) {
  const target = path.resolve(baseDir, command);

  if (!isInsideBase(target)) {
    throw new Error("Ruta bloqueada por seguridad: fuera del proyecto HMA.");
  }

  return target;
}

function createCmdWrapper(target) {
  const safeName = `hma_run_${Date.now()}_${Math.floor(Math.random() * 100000)}.cmd`;
  const wrapperPath = path.join(runtimeDir, safeName);

  const content = [
    "@echo off",
    `cd /d "${baseDir}"`,
    "set HMA_DASHBOARD_MODE=1",
    "set HMA_NO_NOTEPAD=1",
    "set HMA_NO_OPEN_TXT=1",
    `call "${target}"`,
    "exit /b %ERRORLEVEL%"
  ].join("\r\n");

  fs.writeFileSync(wrapperPath, content, { encoding: "utf8" });

  return wrapperPath;
}

async function runBatOrCmd(target, startMs, title, command) {
  const wrapperPath = createCmdWrapper(target);

  const result = await runProcess("cmd.exe", ["/d", "/c", wrapperPath], baseDir);

  try {
    fs.unlinkSync(wrapperPath);
  } catch {}

  return appendGeneratedFilesToResult({
    ...result,
    title,
    command,
    stdout: result.stdout || `Comando ejecutado:\n${target}`
  }, startMs);
}

async function executeAction(payload) {
  const startMs = Date.now();
  const title = payload?.title || "Accion HMA";
  const command = payload?.command || "";

  if (!command) {
    return {
      ok: false,
      title,
      command,
      stdout: "",
      stderr: "No se recibio comando para ejecutar."
    };
  }

  if (command.startsWith("set-report-frequency:")) {
    const frequency = command.replace("set-report-frequency:", "").trim();
    const allowed = ["1h", "3h", "5h", "7h", "12h", "1d", "2d", "1w"];

    if (!allowed.includes(frequency)) {
      return {
        ok: false,
        title,
        command,
        stdout: "",
        stderr: `Frecuencia no permitida: ${frequency}`
      };
    }

    const script = path.join(baseDir, "scripts", "setup_report_frequency_task.ps1");

    const result = await runProcess("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      script,
      "-Frequency",
      frequency
    ], baseDir);

    return appendGeneratedFilesToResult({
      ...result,
      title,
      command,
      stdout: result.stdout || `Frecuencia configurada: ${frequency}`
    }, startMs);
  }

  if (command === "git-status") {
    const result = await runProcess("git", ["-C", baseDir, "status", "--short"]);

    return {
      ...result,
      title,
      command,
      stdout: result.stdout.trim() || "Git limpio."
    };
  }

  if (command === "clients-list") {
    const psCommand = `
      $root = Join-Path '${baseDir.replace(/'/g, "''")}' 'clientes';
      if (!(Test-Path $root)) {
        Write-Host 'No existe carpeta clientes.';
        exit 0;
      }
      Get-ChildItem $root -Directory |
      Where-Object { $_.Name -ne '_template' } |
      Select-Object Name, FullName |
      Format-Table -AutoSize
    `;

    const result = await runProcess("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-Command",
      psCommand
    ]);

    return appendGeneratedFilesToResult({
      ...result,
      title,
      command
    }, startMs);
  }

  const target = resolveSafeTarget(command);

  if (!fs.existsSync(target)) {
    return {
      ok: false,
      title,
      command,
      stdout: "",
      stderr: `No existe la ruta:\n${target}`
    };
  }

  const stat = fs.statSync(target);

  if (stat.isDirectory()) {
    const openResult = await shell.openPath(target);

    return {
      ok: !openResult,
      title,
      command,
      stdout: openResult ? "" : `Carpeta abierta:\n${target}`,
      stderr: openResult || ""
    };
  }

  const ext = path.extname(target).toLowerCase();

  if (ext === ".bat" || ext === ".cmd") {
    return await runBatOrCmd(target, startMs, title, command);
  }

  if (ext === ".ps1") {
    const result = await runProcess("powershell.exe", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      target
    ], baseDir);

    return appendGeneratedFilesToResult({
      ...result,
      title,
      command,
      stdout: result.stdout || `Script ejecutado:\n${target}`
    }, startMs);
  }

  if ([".xlsx", ".xls", ".csv"].includes(ext)) {
    const openResult = await shell.openPath(target);

    return {
      ok: !openResult,
      title,
      command,
      stdout: openResult ? "" : `Archivo abierto:\n${target}`,
      stderr: openResult || ""
    };
  }

  if ([".txt", ".log"].includes(ext)) {
    return {
      ok: true,
      title,
      command,
      stdout: `Archivo leido dentro del dashboard:\n${target}\n\n${readTextFileSafe(target)}`,
      stderr: ""
    };
  }

  return {
    ok: false,
    title,
    command,
    stdout: "",
    stderr: `Extension no soportada: ${ext}`
  };
}

ipcMain.handle("hma:run-action", async (_event, payload) => {
  try {
    return await executeAction(payload);
  } catch (error) {
    return {
      ok: false,
      title: payload?.title || "Accion HMA",
      command: payload?.command || "",
      stdout: "",
      stderr: error?.message || String(error)
    };
  }
});

ipcMain.handle("hma:get-base-dir", async () => {
  return baseDir;
});

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
