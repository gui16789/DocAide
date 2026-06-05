const fs = require("fs");
const path = require("path");
const {
  businessRuleLeaves,
  expandCoverageEntries,
  loadJson
} = require("../lib/rule-coverage");

const ROOT = path.resolve(__dirname, "..");
const DOC_PATH = path.join(ROOT, "docs", "rule-execution-matrix.md");

function ensureDir(filePath) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function escapeCell(value) {
  return String(value ?? "")
    .replace(/\r?\n/g, " ")
    .replace(/\|/g, "\\|");
}

function formatApplies(entry) {
  return (entry.applies || [])
    .map((item) => `${item.symbol} (${item.file})`)
    .join("<br>");
}

function formatVerification(entry) {
  return (entry.verification || [])
    .map((item) => {
      if (item.type === "document-check") return item.label || `包含：${item.labelIncludes}`;
      if (item.type === "unit-test") return `测试：${item.name}`;
      if (item.type === "code-path") return `代码路径：${item.text}`;
      return `${item.type}`;
    })
    .join("<br>");
}

function render() {
  const rules = loadJson(ROOT, "data/rules.json");
  const coverage = loadJson(ROOT, "data/rule-coverage.json");
  const leaves = businessRuleLeaves(rules).sort();
  const { byPath, unmatchedPatterns } = expandCoverageEntries(coverage, leaves);
  const missing = leaves.filter((leaf) => !byPath.has(leaf));

  const lines = [];
  lines.push("# 规则执行矩阵");
  lines.push("");
  lines.push("本文件由 `node scripts/generate-rule-coverage-doc.js` 根据 `data/rule-coverage.json` 生成。");
  lines.push("");
  lines.push(`- 当前业务规则叶子项：${leaves.length}`);
  lines.push(`- 覆盖条目：${coverage.entries.length}`);
  lines.push(`- 未覆盖规则：${missing.length}`);
  lines.push(`- 未匹配规则模式：${unmatchedPatterns.length}`);
  lines.push("");
  lines.push("## 覆盖分组");
  lines.push("");
  lines.push("| 分组 | 风险 | 规则路径 | 应用位置 | 校验方式 | 备注 |");
  lines.push("| --- | --- | --- | --- | --- | --- |");
  for (const entry of coverage.entries) {
    lines.push([
      entry.title || entry.id,
      entry.risk || "",
      (entry.paths || []).join("<br>"),
      formatApplies(entry),
      formatVerification(entry),
      entry.notes || ""
    ].map(escapeCell).join(" | ").replace(/^/, "| ").replace(/$/, " |"));
  }

  lines.push("");
  lines.push("## 规则展开");
  lines.push("");
  lines.push("| 规则路径 | 覆盖分组 | 风险 |");
  lines.push("| --- | --- | --- |");
  for (const leaf of leaves) {
    const entries = byPath.get(leaf) || [];
    lines.push(`| ${escapeCell(leaf)} | ${escapeCell(entries.map((entry) => entry.title || entry.id).join("<br>"))} | ${escapeCell(entries.map((entry) => entry.risk || "").join("<br>"))} |`);
  }

  if (missing.length || unmatchedPatterns.length) {
    lines.push("");
    lines.push("## 待处理");
    lines.push("");
    for (const leaf of missing) lines.push(`- 未覆盖规则：\`${leaf}\``);
    for (const item of unmatchedPatterns) lines.push(`- 未匹配模式：\`${item.entry}:${item.pattern}\``);
  }

  lines.push("");
  ensureDir(DOC_PATH);
  fs.writeFileSync(DOC_PATH, `${lines.join("\n")}\n`, "utf8");
}

render();
