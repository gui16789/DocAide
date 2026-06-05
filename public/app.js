const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

const state = {
  rules: null,
  ruleVersions: [],
  latestResult: null
};

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
  $("#signer").value = "阮江";
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
    time: new Date().toISOString(),
    originalName: result.originalName,
    pages: result.pages,
    downloads: result.downloads
  });
  localStorage.setItem("redhead-history", JSON.stringify(history.slice(0, 12)));
  renderHistory();
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

async function processDocument(event) {
  event.preventDefault();
  const input = $("#wordFile");
  if (!input.files.length) {
    setStatus("请选择 Word 文件", [{ Label: "未选择文件", Status: "warn" }]);
    return;
  }

  const formData = new FormData();
  formData.append("file", input.files[0]);
  formData.append("meta", JSON.stringify(buildMeta()));

  setBusy(true, "Word 正在套红并导出 PDF");
  setStatus("处理中", [{ Label: "任务已提交到本地 Word", Status: "pass" }]);

  try {
    const response = await fetch("/api/process", {
      method: "POST",
      body: formData
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || result.ok === false) {
      const error = new Error(result.error || "处理失败");
      error.result = result;
      throw error;
    }
    state.latestResult = result;
    renderPreview(result);
    setStatus(result.validationPassed === false ? "处理完成，存在未通过校验" : "处理完成", result.checks || []);
    saveHistoryItem(result);
  } catch (error) {
    setStatus(error.message, buildFailureChecks(error.result));
  } finally {
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
  $("#wordFile").addEventListener("change", (event) => {
    const file = event.target.files[0];
    $("#fileName").textContent = file ? file.name : "选择 Word 文件";
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
}

bindEvents();
renderHistory();
loadRules().catch((error) => {
  setStatus(error.message, [{ Label: "规则加载失败", Status: "fail" }]);
});
