const http = require("http");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const crypto = require("crypto");
const { spawn, spawnSync } = require("child_process");
const { assertValidRules, loadSchema } = require("./lib/rules-validator");

const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "public");
const DATA_DIR = path.join(ROOT, "data");
const UPLOAD_DIR = path.join(ROOT, "uploads");
const OUTPUT_DIR = path.join(ROOT, "output");
const SCRIPTS_DIR = path.join(ROOT, "scripts");
const DEFAULT_RULES = path.join(DATA_DIR, "default-rules.json");
const RULES_PATH = path.join(DATA_DIR, "rules.json");
const RULES_HISTORY_PATH = path.join(DATA_DIR, "rules-history.json");
const RULES_SCHEMA_PATH = path.join(DATA_DIR, "rules.schema.json");
const JOB_TIMEOUT_MS = Number(process.env.REDHEAD_JOB_TIMEOUT_MS || 300000);
const RULES_SCHEMA = loadSchema(RULES_SCHEMA_PATH);

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".log": "text/plain; charset=utf-8",
  ".pdf": "application/pdf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".doc": "application/msword",
  ".svg": "image/svg+xml"
};

let queue = Promise.resolve();

function getPort() {
  const index = process.argv.indexOf("--port");
  if (index >= 0 && process.argv[index + 1]) return Number(process.argv[index + 1]);
  return Number(process.env.PORT || 3721);
}

function getHost() {
  const index = process.argv.indexOf("--host");
  if (index >= 0 && process.argv[index + 1]) return process.argv[index + 1];
  return process.env.HOST || "0.0.0.0";
}

async function ensureDirs() {
  await Promise.all([DATA_DIR, UPLOAD_DIR, OUTPUT_DIR, path.join(ROOT, "logs")].map((dir) => fsp.mkdir(dir, { recursive: true })));
  if (!fs.existsSync(RULES_PATH)) {
    const defaults = normalizeRulesPayload(JSON.parse(await fsp.readFile(DEFAULT_RULES, "utf8")));
    assertValidRules(defaults, RULES_SCHEMA, "默认规则");
    const initial = stampRules(defaults, "初始化默认规则");
    assertValidRules(initial, RULES_SCHEMA, "初始化规则");
    await fsp.writeFile(RULES_PATH, `${JSON.stringify(initial, null, 2)}\n`, "utf8");
  } else {
    const rawCurrent = JSON.parse(await fsp.readFile(RULES_PATH, "utf8"));
    const current = normalizeRulesPayload(rawCurrent);
    assertValidRules(current, RULES_SCHEMA, "当前规则");
    if (JSON.stringify(current) !== JSON.stringify(rawCurrent)) {
      const migrated = stampRules(current, "迁移旧规则字段");
      migrated.ruleId = current.ruleId || migrated.ruleId;
      assertValidRules(migrated, RULES_SCHEMA, "当前规则");
      await fsp.writeFile(RULES_PATH, `${JSON.stringify(migrated, null, 2)}\n`, "utf8");
    } else if (!current.ruleId || !current.revisionId) {
      const stamped = stampRules(current, "补齐规则版本信息");
      assertValidRules(stamped, RULES_SCHEMA, "当前规则");
      await fsp.writeFile(RULES_PATH, `${JSON.stringify(stamped, null, 2)}\n`, "utf8");
    }
  }
  if (!fs.existsSync(RULES_HISTORY_PATH)) {
    const current = JSON.parse(await fsp.readFile(RULES_PATH, "utf8"));
    await fsp.writeFile(RULES_HISTORY_PATH, `${JSON.stringify([toRuleHistoryEntry(current, "初始化默认规则")], null, 2)}\n`, "utf8");
  }
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload, null, 2);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  res.end(body);
}

function sendText(res, status, text) {
  res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(text);
}

function readRequestBody(req, maxBytes = 200 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > maxBytes) {
        reject(new Error("上传文件超过 200MB 限制"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function sanitizeFileName(name) {
  const fallback = `upload-${Date.now()}`;
  const base = path.basename(String(name || fallback));
  return base.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").slice(0, 180) || fallback;
}

function splitBuffer(buffer, delimiter) {
  const parts = [];
  let start = 0;
  let index = buffer.indexOf(delimiter, start);
  while (index !== -1) {
    parts.push(buffer.subarray(start, index));
    start = index + delimiter.length;
    index = buffer.indexOf(delimiter, start);
  }
  parts.push(buffer.subarray(start));
  return parts;
}

function parseContentDisposition(header) {
  const result = {};
  const match = header.match(/content-disposition:\s*form-data;(.+)/i);
  if (!match) return result;
  for (const segment of match[1].split(";")) {
    const [key, rawValue] = segment.trim().split("=");
    if (!key || rawValue === undefined) continue;
    result[key.toLowerCase()] = rawValue.replace(/^"|"$/g, "");
  }
  return result;
}

function parseMultipart(body, contentType) {
  const boundaryMatch = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/i);
  if (!boundaryMatch) throw new Error("缺少 multipart boundary");
  const boundary = Buffer.from(`--${boundaryMatch[1] || boundaryMatch[2]}`);
  const rawParts = splitBuffer(body, boundary);
  const fields = {};
  const files = {};

  for (let part of rawParts) {
    if (part.length === 0) continue;
    if (part.subarray(0, 2).toString() === "--") continue;
    if (part.subarray(0, 2).toString() === "\r\n") part = part.subarray(2);
    if (part.subarray(part.length - 2).toString() === "\r\n") part = part.subarray(0, part.length - 2);

    const headerEnd = part.indexOf(Buffer.from("\r\n\r\n"));
    if (headerEnd < 0) continue;
    const headerText = part.subarray(0, headerEnd).toString("utf8");
    const content = part.subarray(headerEnd + 4);
    const disposition = parseContentDisposition(headerText);
    if (!disposition.name) continue;

    if (disposition.filename !== undefined) {
      files[disposition.name] = {
        filename: sanitizeFileName(disposition.filename),
        data: content,
        contentType: (headerText.match(/content-type:\s*([^\r\n]+)/i) || [])[1] || "application/octet-stream"
      };
    } else {
      fields[disposition.name] = content.toString("utf8");
    }
  }

  return { fields, files };
}

async function loadRules() {
  const raw = await fsp.readFile(RULES_PATH, "utf8");
  return JSON.parse(raw);
}

function normalizeRulesPayload(payload) {
  const rules = JSON.parse(JSON.stringify(payload || {}));
  if (rules.pageNumber && Object.prototype.hasOwnProperty.call(rules.pageNumber, "bodyOnly")) {
    delete rules.pageNumber.bodyOnly;
  }
  return rules;
}

async function saveRules(payload) {
  const rules = stampRules(normalizeRulesPayload(payload), "手动保存");
  assertValidRules(rules, RULES_SCHEMA, "保存规则");
  await fsp.writeFile(RULES_PATH, `${JSON.stringify(rules, null, 2)}\n`, "utf8");
  await appendRuleHistory(rules, "手动保存");
  return rules;
}

function stampRules(payload, note = "") {
  return {
    ...payload,
    ruleId: payload.ruleId || crypto.randomUUID(),
    revisionId: crypto.randomUUID(),
    updatedAt: new Date().toISOString(),
    note
  };
}

function toRuleHistoryEntry(rules, note = "") {
  return {
    id: rules.revisionId || crypto.randomUUID(),
    ruleId: rules.ruleId || crypto.randomUUID(),
    version: rules.version || "未命名规则",
    updatedAt: rules.updatedAt || new Date().toISOString(),
    note: note || rules.note || "",
    rules
  };
}

async function loadRuleHistory() {
  try {
    return JSON.parse(await fsp.readFile(RULES_HISTORY_PATH, "utf8"));
  } catch {
    return [];
  }
}

async function appendRuleHistory(rules, note) {
  const history = await loadRuleHistory();
  history.unshift(toRuleHistoryEntry(rules, note));
  await fsp.writeFile(RULES_HISTORY_PATH, `${JSON.stringify(history.slice(0, 50), null, 2)}\n`, "utf8");
}

async function restoreRuleVersion(id) {
  const history = await loadRuleHistory();
  const entry = history.find((item) => item.id === id);
  if (!entry) throw new Error("未找到规则版本");
  const restored = stampRules(normalizeRulesPayload(entry.rules), `回滚自 ${entry.id}`);
  restored.ruleId = entry.rules.ruleId || restored.ruleId;
  restored.restoredFrom = entry.id;
  assertValidRules(restored, RULES_SCHEMA, "回滚规则");
  await fsp.writeFile(RULES_PATH, `${JSON.stringify(restored, null, 2)}\n`, "utf8");
  await appendRuleHistory(restored, `回滚自 ${entry.id}`);
  return restored;
}

function enqueue(task) {
  const run = queue.then(task, task);
  queue = run.catch(() => {});
  return run;
}

function readJsonIfExists(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch {
    return null;
  }
}

async function writeJobStatus(statusPath, status) {
  if (!statusPath) return;
  const payload = {
    ...status,
    updatedAt: new Date().toISOString()
  };
  await fsp.writeFile(statusPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

function describeJobStatus(status) {
  if (!status) return "";
  const parts = [];
  if (status.stage) parts.push(`最后阶段：${status.stage}`);
  if (status.detail) parts.push(status.detail);
  if (Number.isFinite(Number(status.elapsedMs))) {
    parts.push(`已耗时 ${Math.round(Number(status.elapsedMs) / 1000)} 秒`);
  }
  if (status.wordPid) parts.push(`Word PID ${status.wordPid}`);
  return parts.join("，");
}

function runPowerShell(args, cwd, options = {}) {
  return new Promise((resolve, reject) => {
    const executable = getPowerShellExecutable();
    const startedAt = new Date();
    const child = spawn(executable, ["-NoProfile", "-ExecutionPolicy", "Bypass", ...args], {
      cwd,
      windowsHide: true
    });
    let stdout = "";
    let stderr = "";
    let timedOut = false;
    const timeoutMs = options.timeoutMs || JOB_TIMEOUT_MS;
    const timer = setTimeout(() => {
      timedOut = true;
      cleanupTimedOutWordJob(child.pid, options, startedAt);
      setTimeout(() => cleanupTimedOutWordJob(null, options, startedAt), 3000).unref?.();
    }, timeoutMs);
    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });
    child.on("error", (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.on("close", (code) => {
      clearTimeout(timer);
      const status = readJsonIfExists(options.statusPath);
      const diagnostics = {
        startedAt: startedAt.toISOString(),
        finishedAt: new Date().toISOString(),
        exitCode: code,
        timedOut,
        timeoutMs,
        status
      };
      writeProcessLog(options.logPath, stdout, stderr, diagnostics);
      if (timedOut) {
        cleanupTimedOutWordJob(null, options, startedAt);
        const statusDetail = describeJobStatus(status);
        const suffix = statusDetail ? `。${statusDetail}` : "";
        const error = new Error(`处理超时，已终止本次 Word 任务（${Math.round(timeoutMs / 1000)} 秒）${suffix}`);
        error.stdout = stdout;
        error.stderr = stderr;
        error.status = status;
        error.diagnostics = diagnostics;
        reject(error);
        return;
      }
      if (code !== 0) {
        const error = new Error(stderr || stdout || `PowerShell 退出码 ${code}`);
        error.stdout = stdout;
        error.stderr = stderr;
        error.status = status;
        error.diagnostics = diagnostics;
        reject(error);
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

function writeProcessLog(logPath, stdout, stderr, diagnostics = null) {
  if (!logPath) return;
  const sections = [
    `[diagnostics]`,
    diagnostics ? JSON.stringify(diagnostics, null, 2) : "",
    `[stdout]`,
    stdout || "",
    `[stderr]`,
    stderr || ""
  ];
  const text = sections.join("\n");
  try {
    fs.writeFileSync(logPath, text, "utf8");
  } catch {
    // Logging should not mask the real processing result.
  }
}

function killProcessTree(pid) {
  if (!pid) return;
  spawnSync("taskkill.exe", ["/PID", String(pid), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
}

function cleanupTimedOutWordJob(childPid, options, startedAt) {
  killProcessTree(childPid);
  cleanupWordProcess(options.pidPath);
  cleanupWordProcessFromStatus(options.statusPath);
  cleanupRecentWordProcesses(startedAt);
}

function cleanupWordProcess(pidPath) {
  if (!pidPath || !fs.existsSync(pidPath)) return;
  try {
    const info = JSON.parse(fs.readFileSync(pidPath, "utf8"));
    if (info.wordPid) {
      spawnSync("taskkill.exe", ["/PID", String(info.wordPid), "/F"], { windowsHide: true, stdio: "ignore" });
    }
  } catch {
    // Best-effort cleanup; the job result will still report the timeout.
  }
}

function cleanupWordProcessFromStatus(statusPath) {
  const status = readJsonIfExists(statusPath);
  if (!status || !status.wordPid) return;
  spawnSync("taskkill.exe", ["/PID", String(status.wordPid), "/F"], { windowsHide: true, stdio: "ignore" });
}

function cleanupRecentWordProcesses(startedAt) {
  const cutoff = new Date(startedAt.getTime() - 5000).toISOString();
  const command = [
    `$cutoff = [datetime]::Parse('${cutoff}')`,
    `Get-Process WINWORD -ErrorAction SilentlyContinue |`,
    `Where-Object { $_.StartTime -ge $cutoff } |`,
    `Stop-Process -Force -ErrorAction SilentlyContinue`
  ].join(" ");
  spawnSync(getPowerShellExecutable(), ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command], {
    windowsHide: true,
    stdio: "ignore"
  });
}

function getPowerShellExecutable() {
  const candidates = [
    process.env.PWSH_PATH,
    "C:\\Program Files\\PowerShell\\7\\pwsh.exe",
    "pwsh.exe",
    "powershell.exe"
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (candidate.includes(":\\") && fs.existsSync(candidate)) return candidate;
    if (!candidate.includes(":\\") && commandExists(candidate)) return candidate;
  }
  return "powershell.exe";
}

function commandExists(command) {
  const result = spawnSync("where.exe", [command], { windowsHide: true, stdio: "ignore" });
  return result.status === 0;
}

async function processDocument(body, contentType) {
  const { fields, files } = parseMultipart(body, contentType);
  const upload = files.file;
  if (!upload) throw new Error("没有收到 Word 文件");

  const ext = path.extname(upload.filename).toLowerCase();
  if (![".doc", ".docx"].includes(ext)) throw new Error("仅支持 .doc 和 .docx 文件");

  const jobId = `${new Date().toISOString().replace(/[-:.TZ]/g, "")}-${crypto.randomBytes(3).toString("hex")}`;
  const uploadJobDir = path.join(UPLOAD_DIR, jobId);
  const outputJobDir = path.join(OUTPUT_DIR, jobId);
  await fsp.mkdir(uploadJobDir, { recursive: true });
  await fsp.mkdir(outputJobDir, { recursive: true });

  const inputPath = path.join(uploadJobDir, upload.filename);
  const rulesSnapshotPath = path.join(outputJobDir, "rules.snapshot.json");
  const metaPath = path.join(outputJobDir, "meta.json");
  const resultPath = path.join(outputJobDir, "result.json");
  const pidPath = path.join(outputJobDir, "process-pids.json");
  const processLogPath = path.join(outputJobDir, "process.log");
  const statusPath = path.join(outputJobDir, "job-status.json");

  const meta = fields.meta ? JSON.parse(fields.meta) : {};
  const rules = await loadRules();
  await fsp.writeFile(inputPath, upload.data);
  await fsp.writeFile(rulesSnapshotPath, `${JSON.stringify(rules, null, 2)}\n`, "utf8");
  await fsp.writeFile(metaPath, `${JSON.stringify(meta, null, 2)}\n`, "utf8");
  await writeJobStatus(statusPath, {
    stage: "powershell-starting",
    detail: "正在启动 PowerShell 套红脚本",
    elapsedMs: 0,
    powershellPid: null,
    wordPid: null,
    jobId,
    originalName: upload.filename
  });

  await runPowerShell([
    "-File",
    path.join(SCRIPTS_DIR, "redhead.ps1"),
    "-InputPath",
    inputPath,
    "-OutputDir",
    outputJobDir,
    "-RulesPath",
    rulesSnapshotPath,
    "-MetaPath",
    metaPath,
    "-ResultPath",
    resultPath,
    "-PidPath",
    pidPath,
    "-StatusPath",
    statusPath
  ], ROOT, { timeoutMs: JOB_TIMEOUT_MS, pidPath, logPath: processLogPath, statusPath }).catch((error) => {
    error.jobId = jobId;
    error.status = error.status || readJsonIfExists(statusPath);
    error.downloads = {
      log: `/jobs/${jobId}/${encodeURIComponent(path.basename(processLogPath))}`,
      status: `/jobs/${jobId}/${encodeURIComponent(path.basename(statusPath))}`
    };
    throw error;
  });

  const result = JSON.parse(await fsp.readFile(resultPath, "utf8"));
  result.jobId = jobId;
  result.originalName = upload.filename;
  result.rule = {
    version: rules.version,
    ruleId: rules.ruleId,
    revisionId: rules.revisionId,
    updatedAt: rules.updatedAt
  };
  result.downloads = {
    docx: `/jobs/${jobId}/${encodeURIComponent(path.basename(result.outputDocx))}`,
    pdf: `/jobs/${jobId}/${encodeURIComponent(path.basename(result.outputPdf))}`,
    log: `/jobs/${jobId}/${encodeURIComponent(path.basename(processLogPath))}`,
    status: `/jobs/${jobId}/${encodeURIComponent(path.basename(statusPath))}`
  };
  result.status = readJsonIfExists(statusPath);
  return result;
}

async function serveStatic(req, res, pathname) {
  const filePath = pathname === "/" ? path.join(PUBLIC_DIR, "index.html") : path.join(PUBLIC_DIR, pathname);
  const resolved = path.resolve(filePath);
  if (!resolved.startsWith(path.resolve(PUBLIC_DIR))) {
    sendText(res, 403, "Forbidden");
    return;
  }
  try {
    const data = await fsp.readFile(resolved);
    const type = MIME_TYPES[path.extname(resolved).toLowerCase()] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": type,
      "Cache-Control": "no-store"
    });
    res.end(data);
  } catch {
    sendText(res, 404, "Not found");
  }
}

async function serveJobFile(req, res, pathname) {
  const match = pathname.match(/^\/jobs\/([^/]+)\/(.+)$/);
  if (!match) {
    sendText(res, 404, "Not found");
    return;
  }
  const jobId = sanitizeFileName(decodeURIComponent(match[1]));
  const fileName = sanitizeFileName(decodeURIComponent(match[2]));
  const filePath = path.resolve(OUTPUT_DIR, jobId, fileName);
  if (!filePath.startsWith(path.resolve(OUTPUT_DIR))) {
    sendText(res, 403, "Forbidden");
    return;
  }
  try {
    const data = await fsp.readFile(filePath);
    const type = MIME_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream";
    const disposition = type.includes("pdf") ? "inline" : "attachment";
    res.writeHead(200, {
      "Content-Type": type,
      "Content-Disposition": `${disposition}; filename*=UTF-8''${encodeURIComponent(fileName)}`,
      "Cache-Control": "no-store"
    });
    res.end(data);
  } catch {
    sendText(res, 404, "Not found");
  }
}

async function handle(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const pathname = decodeURIComponent(url.pathname);

  try {
    if (req.method === "GET" && pathname === "/api/rules") {
      sendJson(res, 200, await loadRules());
      return;
    }

    if (req.method === "PUT" && pathname === "/api/rules") {
      const body = await readRequestBody(req, 2 * 1024 * 1024);
      sendJson(res, 200, await saveRules(JSON.parse(body.toString("utf8"))));
      return;
    }

    if (req.method === "GET" && pathname === "/api/rules/versions") {
      const history = await loadRuleHistory();
      sendJson(res, 200, history.map(({ id, ruleId, version, updatedAt, note }) => ({ id, ruleId, version, updatedAt, note })));
      return;
    }

    if (req.method === "POST" && pathname === "/api/rules/restore") {
      const body = await readRequestBody(req, 2 * 1024 * 1024);
      const { id } = JSON.parse(body.toString("utf8"));
      sendJson(res, 200, await restoreRuleVersion(id));
      return;
    }

    if (req.method === "POST" && pathname === "/api/process") {
      const body = await readRequestBody(req);
      const result = await enqueue(() => processDocument(body, req.headers["content-type"] || ""));
      sendJson(res, 200, result);
      return;
    }

    if (req.method === "GET" && pathname.startsWith("/jobs/")) {
      await serveJobFile(req, res, pathname);
      return;
    }

    if (req.method === "GET") {
      await serveStatic(req, res, pathname);
      return;
    }

    sendText(res, 405, "Method not allowed");
  } catch (error) {
    sendJson(res, error.statusCode || 500, {
      error: error.message,
      details: error.details || error.stderr || error.stdout || undefined,
      jobId: error.jobId || undefined,
      status: error.status || undefined,
      diagnostics: error.diagnostics || undefined,
      downloads: error.downloads || undefined
    });
  }
}

ensureDirs()
  .then(() => {
    const port = getPort();
    const host = getHost();
    http.createServer(handle).listen(port, host, () => {
      console.log(`Redhead web tool listening on http://${host}:${port}`);
    });
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });
