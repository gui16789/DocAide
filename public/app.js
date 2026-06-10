const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

const state = {
  rules: null,
  ruleVersions: [],
  latestResult: null,
  selectedFile: null
};

const WORD_FILE_PATTERN = /\.(doc|docx)$/i;

function getByPath(object, path) {
  return path.split(".").reduce((value, key) => (value == null ? undefined : value[key]), object);
}

function setByPath(object, path, value) {
  const parts = path.split(".");
  let target = object;
  for (let i = 0; i < parts.length - 1; i++) {
    target[parts[i]] ||= {};
    target = target[parts[i]];
  }
  target[parts.at(-1)] = value;
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function formatStamp(rules) {
  const updated = rules.updatedAt ? new Date(rules.updatedAt) : null;
  const date = updated && !Number.isNaN(updated.valueOf()) ? updated.toLocaleString("zh-CN") : rules.updatedAt;
  return `${rules.version || "未命名规则"} · ${date || "未标记日期"}`;
}

function fillRuleInputs() {
  $$("[data-rule]").forEach((input) => {
    const value = getByPath(state.rules, input.dataset.rule);
    if (input.type === "checkbox") {
      input.checked = Boolean(value);
    } else {
      input.value = value ?? "";
    }
  });

  $("#documentNo").value = getByPath(state.rules, "documentNo.defaultText") || "";
  $("#company").value = getByPath(state.rules, "signature.company") || "";
  $("#date").value = getByPath(state.rules, "signature.date") || "";
  $("#ruleStamp").textContent = formatStamp(state.rules);
  renderRuleVersions();
}

function collectRules() {
  const next = clone(state.rules);
  $$("[data-rule]").forEach((input) => {
    let value;
    if (input.type === "checkbox") {
      value = input.checked;
    } else if (input.type === "number") {
      value = Number(input.value);
    } else {
      value = input.value;
    }
    setByPath(next, input.dataset.rule, value);
  });
  return next;
}

async function loadRules() {
  const [rulesResponse, versionsResponse] = await Promise.all([
    fetch("/api/rules"),
    fetch("/api/rules/versions")
  ]);
  if (!rulesResponse.ok) throw new Error("规则加载失败");
  if (!versionsResponse.ok) throw new Error("规则版本加载失败");
  state.rules = await rulesResponse.json();
  state.ruleVersions = await versionsResponse.json();
  fillRuleInputs();
}

async function saveRules() {
  const next = collectRules();
  setBusy(true, "保存规则中");
  try {
    const response = await fetch("/api/rules", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(next)
    });
    if (!response.ok) throw new Error("规则保存失败");
    state.rules = await response.json();
    await loadRuleVersions();
    fillRuleInputs();
    setStatus("规则已保存", []);
  } finally {
    setBusy(false);
  }
}

async function loadRuleVersions() {
  const response = await fetch("/api/rules/versions");
  if (!response.ok) throw new Error("规则版本加载失败");
  state.ruleVersions = await response.json();
  renderRuleVersions();
}

function renderRuleVersions() {
  const select = $("#ruleVersionSelect");
  if (!select) return;
  select.innerHTML = "";
  if (!state.ruleVersions.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "暂无历史版本";
    select.append(option);
    return;
  }
  state.ruleVersions.forEach((version) => {
    const option = document.createElement("option");
    option.value = version.id;
    const time = version.updatedAt ? new Date(version.updatedAt).toLocaleString("zh-CN") : "未标记时间";
    option.textContent = `${time} · ${version.version || "未命名规则"} · ${version.note || "保存"}`;
    select.append(option);
  });
}

async function restoreRuleVersion() {
  const id = $("#ruleVersionSelect").value;
  if (!id) {
    setStatus("请选择规则版本", [{ Label: "未选择历史版本", Status: "warn" }]);
    return;
  }
  setBusy(true, "正在回滚规则");
  try {
    const response = await fetch("/api/rules/restore", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id })
    });
    if (!response.ok) throw new Error("规则回滚失败");
    state.rules = await response.json();
    await loadRuleVersions();
    fillRuleInputs();
    setStatus("规则已回滚", []);
  } finally {
    setBusy(false);
  }
}

function setBusy(isBusy, text = "") {
  $("#processBtn").disabled = isBusy;
  $("#saveRulesBtn").disabled = isBusy;
  $("#saveRulesTopBtn").disabled = isBusy;
  $("#restoreRuleBtn").disabled = isBusy;
  $("#refreshVersionsBtn").disabled = isBusy;
  $("#previewCleanupBtn").disabled = isBusy;
  $("#runCleanupBtn").disabled = isBusy;
  if (text) $("#statusText").textContent = text;
}

function setStatus(text, checks = []) {
  $("#statusText").textContent = text;
  const list = $("#checksList");
  list.innerHTML = "";
  checks.forEach((check) => {
    const item = document.createElement("li");
    const badge = document.createElement("span");
    badge.className = `badge ${check.Status || check.status || "pass"}`;
    badge.textContent = check.Status || check.status || "pass";
    const content = document.createElement("span");
    content.textContent = check.Detail || check.detail ? `${check.Label || check.label || ""}：${check.Detail || check.detail}` : check.Label || check.label || "";
    item.append(content);
    item.append(badge);
    list.append(item);
  });
}

function buildMeta() {
  return {
    documentNo: $("#documentNo").value.trim(),
    signer: $("#signer").value.trim(),
    company: $("#company").value.trim(),
    date: $("#date").value.trim(),
    replaceTitle: $("#replaceTitle").checked,
    titleOverride: $("#titleOverride").value.trim()
  };
}

function renderPreview(result) {
  const stage = $("#previewStage");
  stage.innerHTML = "";
  const iframe = document.createElement("iframe");
  iframe.title = "套红 PDF 预览";
  iframe.src = `${result.downloads.pdf}#view=FitH`;
  stage.append(iframe);

  const removed = result.pages.trailingBlankPagesRemoved ? ` · 已清理尾部空白 ${result.pages.trailingBlankPagesRemoved} 页` : "";
  $("#previewMeta").textContent = `共 ${result.pages.total} 页 · 正文 ${result.pages.body} 页 · 版记第 ${result.pages.imprint} 页${removed}`;
  $("#downloadActions").hidden = false;
  $("#downloadDocx").href = result.downloads.docx;
  $("#downloadPdf").href = result.downloads.pdf;
  $("#downloadLog").href = result.downloads.log;
  $("#downloadDocx").download = "";
  $("#downloadPdf").download = "";
  $("#downloadLog").download = "";
}

function buildFailureChecks(result) {
  const checks = [{ Label: "处理失败", Status: "fail" }];
  if (result?.status) {
    const status = result.status;
    const detail = [status.stage, status.detail].filter(Boolean).join("：");
    checks.push({
      Label: "最后阶段",
      Status: "warn",
      Detail: detail || "未写入阶段状态"
    });
  }
  if (result?.downloads?.log) {
    checks.push({
      Label: "处理日志",
      Status: "warn",
      Detail: result.downloads.log
    });
  }
  if (result?.downloads?.status) {
    checks.push({
      Label: "状态文件",
      Status: "warn",
      Detail: result.downloads.status
    });
  }
  return checks;
}

function loadHistory() {
  try {
    return JSON.parse(localStorage.getItem("redhead-history") || "[]");
  } catch {
    return [];
  }
}

function saveHistoryItem(result) {
  const history = loadHistory();
  history.unshift({
    jobId: result.jobId,
    time: new Date().toISOString(),
    originalName: result.originalName,
    pages: result.pages,
    downloads: result.downloads
  });
  localStorage.setItem("redhead-history", JSON.stringify(history.slice(0, 12)));
  renderHistory();
}

function formatBytes(bytes) {
  const value = Number(bytes || 0);
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  if (value < 1024 * 1024 * 1024) return `${(value / 1024 / 1024).toFixed(1)} MB`;
  return `${(value / 1024 / 1024 / 1024).toFixed(1)} GB`;
}

function isSupportedWordFile(file) {
  return Boolean(file?.name && WORD_FILE_PATTERN.test(file.name));
}

function updateSelectedFile(file) {
  state.selectedFile = file || null;
  $("#fileName").textContent = file ? `${file.name} · ${formatBytes(file.size)}` : "选择 Word 文件";
}

function setInputFiles(input, file) {
  if (!file || typeof DataTransfer === "undefined") return;
  try {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
  } catch {
    // Some browser shells keep file inputs read-only; state.selectedFile still handles submit.
  }
}

function collectCleanupOptions() {
  return {
    retentionDays: Number($("#cleanupRetentionDays").value),
    maxJobs: Number($("#cleanupMaxJobs").value)
  };
}

function renderCleanupSummary(result, prefix = "预估") {
  const outputs = result.targetCounts?.outputs || 0;
  const uploads = result.targetCounts?.uploads || 0;
  const protectedCount = result.protectedNames?.length || 0;
  $("#cleanupSummary").textContent = `${prefix}：输出目录 ${outputs} 个，上传目录 ${uploads} 个，预计释放 ${formatBytes(result.reclaimedBytes)}。已保护运行中任务 ${protectedCount} 个。`;
}

function removeDeletedHistory(result) {
  const deletedJobs = new Set((result.deleted?.outputs || []).map((item) => item.name));
  if (!deletedJobs.size) return;
  const history = loadHistory().filter((item) => !item.jobId || !deletedJobs.has(item.jobId));
  localStorage.setItem("redhead-history", JSON.stringify(history));
  renderHistory();
}

async function previewCleanup() {
  const options = collectCleanupOptions();
  const query = new URLSearchParams({
    retentionDays: String(options.retentionDays),
    maxJobs: String(options.maxJobs)
  });
  setBusy(true, "正在预估清理");
  try {
    const response = await fetch(`/api/cleanup/preview?${query}`);
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "预估清理失败");
    renderCleanupSummary(result, "预估清理");
    setStatus("清理预估完成", [
      { Label: "输出目录", Status: "pass", Detail: `${result.targetCounts.outputs} 个` },
      { Label: "上传目录", Status: "pass", Detail: `${result.targetCounts.uploads} 个` },
      { Label: "预计释放", Status: "pass", Detail: formatBytes(result.reclaimedBytes) }
    ]);
  } finally {
    setBusy(false);
  }
}

async function runCleanup() {
  const options = collectCleanupOptions();
  if (!confirm("确认按当前策略清理输出和上传目录？")) return;
  setBusy(true, "正在清理产物");
  try {
    const response = await fetch("/api/cleanup", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(options)
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "清理失败");
    renderCleanupSummary(result, "已清理");
    removeDeletedHistory(result);
    setStatus("清理完成", [
      { Label: "已清理输出目录", Status: "pass", Detail: `${result.targetCounts.outputs} 个` },
      { Label: "已清理上传目录", Status: "pass", Detail: `${result.targetCounts.uploads} 个` },
      { Label: "释放空间", Status: "pass", Detail: formatBytes(result.reclaimedBytes) }
    ]);
  } finally {
    setBusy(false);
  }
}

function renderHistory() {
  const list = $("#historyList");
  const history = loadHistory();
  list.innerHTML = "";
  if (!history.length) {
    const empty = document.createElement("div");
    empty.className = "history-item";
    empty.textContent = "暂无记录";
    list.append(empty);
    return;
  }
  history.forEach((item) => {
    const row = document.createElement("div");
    row.className = "history-item";
    const title = document.createElement("strong");
    title.textContent = item.originalName;
    const meta = document.createElement("span");
    meta.textContent = `${new Date(item.time).toLocaleString("zh-CN")} · ${item.pages.total} 页`;
    const links = document.createElement("div");
    links.className = "history-links";
    links.innerHTML = `<a href="${item.downloads.docx}">DOCX</a><a href="${item.downloads.pdf}" target="_blank">PDF</a>`;
    row.append(title, meta, links);
    list.append(row);
  });
}

function describeProgress(status) {
  if (!status) return "处理中";
  const stage = status.stage || "处理中";
  const detail = status.detail || "";
  const elapsedSec = Number.isFinite(Number(status.elapsedMs)) ? Math.round(Number(status.elapsedMs) / 1000) : null;
  const tail = elapsedSec !== null ? `（已耗时 ${elapsedSec} 秒）` : "";
  return detail ? `${stage}：${detail}${tail}` : `${stage}${tail}`;
}

function buildProgressChecks(status) {
  const checks = [{ Label: "任务已提交到本地 Word", Status: "pass" }];
  if (status?.stage) {
    checks.push({
      Label: `当前阶段：${status.stage}`,
      Status: "warn",
      Detail: status.detail || ""
    });
  }
  if (status?.wordPid) {
    checks.push({ Label: `Word PID ${status.wordPid}`, Status: "warn" });
  }
  return checks;
}

async function processDocument(event) {
  event.preventDefault();
  const input = $("#wordFile");
  const file = state.selectedFile || input.files[0];
  if (!file) {
    setStatus("请选择 Word 文件", [{ Label: "未选择文件", Status: "warn" }]);
    return;
  }

  const formData = new FormData();
  formData.append("file", file);
  formData.append("meta", JSON.stringify(buildMeta()));

  setBusy(true, "Word 正在套红并导出 PDF");
  setStatus("处理中", [{ Label: "任务已提交到本地 Word", Status: "pass" }]);

  let eventSource = null;
  const closeEvents = () => {
    if (eventSource) {
      eventSource.close();
      eventSource = null;
    }
  };

  try {
    const response = await fetch("/api/process", {
      method: "POST",
      body: formData
    });
    const jobId = response.headers.get("X-Job-Id");
    if (jobId && typeof EventSource !== "undefined") {
      eventSource = new EventSource(`/api/jobs/${encodeURIComponent(jobId)}/events`);
      eventSource.addEventListener("status", (ev) => {
        try {
          const status = JSON.parse(ev.data);
          setBusy(true, describeProgress(status));
          setStatus("处理中", buildProgressChecks(status));
        } catch {
          // 无效负载忽略
        }
      });
      eventSource.addEventListener("done", () => closeEvents());
      eventSource.addEventListener("error", () => closeEvents());
    }
    const result = await response.json().catch(() => ({}));
    closeEvents();
    // 因为 X-Job-Id 已通过 flushHeaders 把状态码定为 200，错误也在 200 通道里返回，
    // 必须同时检查 result.error / result.ok 才能识别失败。
    if (!response.ok || result.ok === false || result.error) {
      const error = new Error(result.error || "处理失败");
      error.result = result;
      throw error;
    }
    state.latestResult = result;
    renderPreview(result);
    setStatus(result.validationPassed === false ? "处理完成，存在未通过校验" : "处理完成", result.checks || []);
    saveHistoryItem(result);
  } catch (error) {
    closeEvents();
    setStatus(error.message, buildFailureChecks(error.result));
  } finally {
    closeEvents();
    setBusy(false);
  }
}

function switchTab(tab) {
  $$(".tab").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
  $$(".tab-pane").forEach((pane) => pane.classList.toggle("active", pane.id === `pane-${tab}`));
}

function bindEvents() {
  $$(".tab").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });
  const fileDrop = $(".file-drop");
  const wordFileInput = $("#wordFile");
  let dragDepth = 0;

  wordFileInput.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file && !isSupportedWordFile(file)) {
      event.target.value = "";
      updateSelectedFile(null);
      setStatus("仅支持 Word 文件", [{ Label: file.name, Status: "fail", Detail: "请上传 .doc 或 .docx 文件" }]);
      return;
    }
    updateSelectedFile(file);
  });

  ["dragenter", "dragover"].forEach((type) => {
    fileDrop.addEventListener(type, (event) => {
      event.preventDefault();
      if (type === "dragenter") dragDepth += 1;
      fileDrop.classList.add("dragging");
      event.dataTransfer.dropEffect = "copy";
    });
  });

  fileDrop.addEventListener("dragleave", (event) => {
    event.preventDefault();
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) fileDrop.classList.remove("dragging");
  });

  fileDrop.addEventListener("drop", (event) => {
    event.preventDefault();
    dragDepth = 0;
    fileDrop.classList.remove("dragging");

    const file = Array.from(event.dataTransfer.files || []).find(isSupportedWordFile);
    if (!file) {
      setStatus("仅支持 Word 文件", [{ Label: "拖拽文件无效", Status: "fail", Detail: "请拖入 .doc 或 .docx 文件" }]);
      return;
    }
    setInputFiles(wordFileInput, file);
    updateSelectedFile(file);
    setStatus("文件已选择", [{ Label: file.name, Status: "pass", Detail: formatBytes(file.size) }]);
  });
  $("#processForm").addEventListener("submit", processDocument);
  $("#saveRulesBtn").addEventListener("click", saveRules);
  $("#saveRulesTopBtn").addEventListener("click", saveRules);
  $("#reloadRulesBtn").addEventListener("click", async () => {
    setBusy(true, "重新加载规则");
    try {
      await loadRules();
      setStatus("规则已重新加载", []);
    } finally {
      setBusy(false);
    }
  });
  $("#restoreRuleBtn").addEventListener("click", restoreRuleVersion);
  $("#refreshVersionsBtn").addEventListener("click", async () => {
    setBusy(true, "刷新规则版本");
    try {
      await loadRuleVersions();
      setStatus("规则版本已刷新", []);
    } finally {
      setBusy(false);
    }
  });
  $("#previewCleanupBtn").addEventListener("click", previewCleanup);
  $("#runCleanupBtn").addEventListener("click", runCleanup);
}

bindEvents();
renderHistory();
loadRules().catch((error) => {
  setStatus(error.message, [{ Label: "规则加载失败", Status: "fail" }]);
});
