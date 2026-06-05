const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const { cleanupJobArtifacts } = require("../lib/job-cleanup");
const { loadSchema, validateRules } = require("../lib/rules-validator");
const {
  businessRuleLeaves,
  expandCoverageEntries
} = require("../lib/rule-coverage");

const ROOT = path.resolve(__dirname, "..");
const REPO_ROOT = path.resolve(ROOT, "..");

const tests = [];

function test(name, fn) {
  tests.push({ name, fn });
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function readText(relativePath, encoding = "utf8") {
  return fs.readFileSync(path.join(ROOT, relativePath), encoding).replace(/^\uFEFF/, "");
}

function readJson(relativePath) {
  return JSON.parse(readText(relativePath));
}

function walkLeaves(value, prefix = "", output = []) {
  if (value && typeof value === "object" && !Array.isArray(value)) {
    for (const [key, child] of Object.entries(value)) {
      walkLeaves(child, prefix ? `${prefix}.${key}` : key, output);
    }
  } else {
    output.push(prefix);
  }
  return output;
}

function collectSchemaLeaves(schema, prefix = "", output = []) {
  if (schema.type === "object") {
    for (const [key, child] of Object.entries(schema.properties || {})) {
      collectSchemaLeaves(child, prefix ? `${prefix}.${key}` : key, output);
    }
  } else {
    output.push(prefix);
  }
  return output;
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    encoding: "utf8",
    windowsHide: true,
    ...options
  });
  if (result.status !== 0) {
    throw new Error([
      `${command} ${args.join(" ")} failed with ${result.status}`,
      result.stdout,
      result.stderr
    ].filter(Boolean).join("\n"));
  }
  return result;
}

function powerShellCommand(command) {
  return run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command]);
}

test("规则 JSON 能解析并符合 Schema", () => {
  const schema = loadSchema(path.join(ROOT, "data", "rules.schema.json"));
  for (const file of ["data/rules.json", "data/default-rules.json"]) {
    const errors = validateRules(readJson(file), schema);
    assert(errors.length === 0, `${file} 规则校验失败：\n${errors.join("\n")}`);
  }
});

test("Schema 覆盖当前规则业务叶子项", () => {
  const rules = readJson("data/rules.json");
  const schema = loadSchema(path.join(ROOT, "data", "rules.schema.json"));
  const ruleLeaves = new Set(walkLeaves(rules).filter((item) => !["ruleId", "revisionId", "note", "restoredFrom"].includes(item)));
  const schemaLeaves = new Set(collectSchemaLeaves(schema));
  const missing = [...ruleLeaves].filter((item) => !schemaLeaves.has(item));
  assert(missing.length === 0, `Schema 未覆盖规则项：${missing.join(", ")}`);
});

test("工作台覆盖全部业务规则项", () => {
  const rules = readJson("data/rules.json");
  const html = readText("public/index.html");
  const uiRules = new Set([...html.matchAll(/data-rule="([^"]+)"/g)].map((match) => match[1]));
  const metadata = new Set(["ruleId", "revisionId", "updatedAt", "note", "restoredFrom"]);
  const missing = walkLeaves(rules).filter((item) => !metadata.has(item) && !uiRules.has(item));
  assert(missing.length === 0, `工作台缺少规则项：${missing.join(", ")}`);
  const ruleLeaves = new Set(walkLeaves(rules).filter((item) => !metadata.has(item)));
  const stale = [...uiRules].filter((item) => !ruleLeaves.has(item));
  assert(stale.length === 0, `工作台存在无效规则项：${stale.join(", ")}`);
});

test("脚本读取的规则路径都存在于 Schema", () => {
  const schema = loadSchema(path.join(ROOT, "data", "rules.schema.json"));
  const schemaLeaves = new Set(collectSchemaLeaves(schema));
  const scriptFiles = [
    "scripts/redhead.ps1",
    "scripts/modules/Redhead.Core.ps1",
    "scripts/modules/Redhead.Runner.ps1"
  ].filter((file) => fs.existsSync(path.join(ROOT, file)));
  const rulePaths = new Set();
  for (const file of scriptFiles) {
    const text = readText(file, "utf8");
    for (const match of text.matchAll(/Get-RuleValue\s+\$(?:Rules|rules)\s+["']([^"']+)["']/g)) {
      rulePaths.add(match[1]);
    }
  }
  const missing = [...rulePaths].filter((item) => !schemaLeaves.has(item));
  assert(missing.length === 0, `脚本读取了未定义规则：${missing.join(", ")}`);
});

test("规则执行矩阵覆盖全部业务规则项", () => {
  const rules = readJson("data/rules.json");
  const coverage = readJson("data/rule-coverage.json");
  const leaves = businessRuleLeaves(rules);
  const { byPath, unmatchedPatterns } = expandCoverageEntries(coverage, leaves);
  const missing = leaves.filter((leaf) => !byPath.has(leaf));
  assert(unmatchedPatterns.length === 0, `规则执行矩阵存在无匹配路径：${unmatchedPatterns.map((item) => `${item.entry}:${item.pattern}`).join(", ")}`);
  assert(missing.length === 0, `规则执行矩阵缺少规则项：${missing.join(", ")}`);

  for (const entry of coverage.entries || []) {
    assert(entry.id && entry.title, "规则执行矩阵条目缺少 id/title");
    assert(Array.isArray(entry.paths) && entry.paths.length > 0, `${entry.id} 缺少 paths`);
    assert(Array.isArray(entry.applies) && entry.applies.length > 0, `${entry.id} 缺少 applies`);
    assert(Array.isArray(entry.verification) && entry.verification.length > 0, `${entry.id} 缺少 verification`);
    if (["high", "medium"].includes(entry.risk)) {
      assert(entry.verification.some((item) => item.type === "document-check"), `${entry.id} 缺少文档实测校验`);
    }
  }
});

test("规则执行矩阵引用的应用和校验都存在", () => {
  const coverage = readJson("data/rule-coverage.json");
  const coreText = readText("scripts/modules/Redhead.Core.ps1");
  const testText = readText("tests/run-tests.js");
  const fileCache = new Map();
  const readProjectFile = (relativePath) => {
    if (!fileCache.has(relativePath)) fileCache.set(relativePath, readText(relativePath));
    return fileCache.get(relativePath);
  };

  for (const entry of coverage.entries || []) {
    for (const item of entry.applies || []) {
      assert(item.file && item.symbol, `${entry.id} 的 applies 缺少 file/symbol`);
      const text = readProjectFile(item.file);
      assert(text.includes(item.symbol), `${entry.id} 引用的应用符号不存在：${item.file} -> ${item.symbol}`);
    }

    for (const item of entry.verification || []) {
      if (item.type === "document-check") {
        const needle = item.label || item.labelIncludes;
        assert(needle, `${entry.id} 的 document-check 缺少 label/labelIncludes`);
        assert(coreText.includes(needle), `${entry.id} 引用的文档校验不存在：${needle}`);
      } else if (item.type === "unit-test") {
        assert(item.name && testText.includes(`test("${item.name}"`), `${entry.id} 引用的单元测试不存在：${item.name}`);
      } else if (item.type === "code-path") {
        assert(item.file && item.text, `${entry.id} 的 code-path 缺少 file/text`);
        assert(readProjectFile(item.file).includes(item.text), `${entry.id} 引用的代码路径不存在：${item.file} -> ${item.text}`);
      } else {
        throw new Error(`${entry.id} 使用了未知 verification 类型：${item.type}`);
      }
    }
  }
});

test("规则执行矩阵 Markdown 已生成并覆盖当前规则", () => {
  const rules = readJson("data/rules.json");
  const leaves = businessRuleLeaves(rules);
  const doc = readText("docs/rule-execution-matrix.md");
  assert(doc.includes(`当前业务规则叶子项：${leaves.length}`), "规则执行矩阵 Markdown 规则数量不是当前值");
  const missing = leaves.filter((leaf) => !doc.includes(`| ${leaf} |`));
  assert(missing.length === 0, `规则执行矩阵 Markdown 缺少规则项：${missing.join(", ")}`);
});

test("前端和服务端 JS 语法通过", () => {
  run(process.execPath, ["--check", "server.js"]);
  run(process.execPath, ["--check", "public/app.js"]);
  run(process.execPath, ["--check", "lib/job-cleanup.js"]);
  run(process.execPath, ["--check", "lib/rule-coverage.js"]);
  run(process.execPath, ["--check", "lib/rules-validator.js"]);
  run(process.execPath, ["--check", "scripts/generate-rule-coverage-doc.js"]);
});

test("任务产物清理策略支持预览和执行", async () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "redhead-cleanup-"));
  const outputDir = path.join(root, "output");
  const uploadDir = path.join(root, "uploads");
  fs.mkdirSync(outputDir, { recursive: true });
  fs.mkdirSync(uploadDir, { recursive: true });
  const now = new Date("2026-06-05T00:00:00.000Z");

  function createJob(baseDir, name, daysAgo) {
    const dir = path.join(baseDir, name);
    fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, "file.txt"), name);
    const date = new Date(now.getTime() - daysAgo * 24 * 60 * 60 * 1000);
    fs.utimesSync(path.join(dir, "file.txt"), date, date);
    fs.utimesSync(dir, date, date);
  }

  try {
    for (const name of ["keep-1", "keep-2", "overflow", "old-job"]) {
      const daysAgo = { "keep-1": 0, "keep-2": 1, overflow: 2, "old-job": 40 }[name];
      createJob(outputDir, name, daysAgo);
      createJob(uploadDir, name, daysAgo);
    }
    createJob(uploadDir, "orphan-upload", 40);

    const policy = { retentionDays: 30, maxJobs: 2, dryRun: true };
    const preview = await cleanupJobArtifacts({ outputDir, uploadDir, policy, now });
    assert(preview.targetCounts.outputs === 2, `预估输出目录数量错误：${preview.targetCounts.outputs}`);
    assert(preview.targetCounts.uploads === 3, `预估上传目录数量错误：${preview.targetCounts.uploads}`);
    assert(fs.existsSync(path.join(outputDir, "old-job")), "dry-run 不应删除目录");

    const cleaned = await cleanupJobArtifacts({ outputDir, uploadDir, policy: { ...policy, dryRun: false }, now });
    assert(cleaned.targetCounts.outputs === 2, "实际清理输出目录数量错误");
    assert(!fs.existsSync(path.join(outputDir, "old-job")), "未清理过期输出目录");
    assert(!fs.existsSync(path.join(outputDir, "overflow")), "未按最大任务数清理输出目录");
    assert(!fs.existsSync(path.join(uploadDir, "old-job")), "未清理同名上传目录");
    assert(!fs.existsSync(path.join(uploadDir, "orphan-upload")), "未清理过期孤立上传目录");
    assert(fs.existsSync(path.join(outputDir, "keep-1")), "误删最新输出目录");
    assert(fs.existsSync(path.join(uploadDir, "keep-2")), "误删保留上传目录");
  } finally {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

test("PowerShell 脚本语法通过", () => {
  const files = [
    "scripts/redhead.ps1",
    "scripts/modules/Redhead.Core.ps1",
    "scripts/modules/Redhead.Runner.ps1"
  ].filter((file) => fs.existsSync(path.join(ROOT, file)));
  const checks = files.map((file) => {
    const absolute = path.join(ROOT, file).replace(/'/g, "''");
    return `$tokens=$null; $errors=$null; [System.Management.Automation.Language.Parser]::ParseFile('${absolute}',[ref]$tokens,[ref]$errors) | Out-Null; if ($errors.Count) { $errors | Format-List *; exit 1 }`;
  }).join("; ");
  powerShellCommand(checks);
});

test("规则校验能拒绝非法规则值", () => {
  const schema = loadSchema(path.join(ROOT, "data", "rules.schema.json"));
  const rules = readJson("data/rules.json");
  rules.page.size = "Letter";
  rules.redHeader.color = "red";
  const errors = validateRules(rules, schema);
  assert(errors.some((item) => item.includes("page.size")), "未拒绝非法纸张规格");
  assert(errors.some((item) => item.includes("redHeader.color")), "未拒绝非法颜色");
});

test("可选 Word COM 回归", () => {
  const enabled = process.argv.includes("--word") || process.env.REDHEAD_RUN_WORD_TESTS === "1";
  if (!enabled) return;
  const fixture = path.join(ROOT, "test-input", "tail-blank-test.docx");
  assert(fs.existsSync(fixture), "缺少 Word 回归样例");
  const output = path.join(ROOT, "output", `qa-auto-${Date.now()}`);
  fs.mkdirSync(output, { recursive: true });
  const metaPath = path.join(output, "meta.json");
  const resultPath = path.join(output, "result.json");
  const statusPath = path.join(output, "job-status.json");
  const pidPath = path.join(output, "process-pids.json");
  fs.writeFileSync(metaPath, JSON.stringify({
    documentNo: "安盟保险〔2026〕142号",
    signer: "阮江",
    company: "安盟财产保险有限公司",
    date: "2026年5月9日",
    replaceTitle: false,
    titleOverride: ""
  }, null, 2));
  powerShellCommand([
    `& '${path.join(ROOT, "scripts", "redhead.ps1").replace(/'/g, "''")}'`,
    `-InputPath '${fixture.replace(/'/g, "''")}'`,
    `-OutputDir '${output.replace(/'/g, "''")}'`,
    `-RulesPath '${path.join(ROOT, "data", "rules.json").replace(/'/g, "''")}'`,
    `-MetaPath '${metaPath.replace(/'/g, "''")}'`,
    `-ResultPath '${resultPath.replace(/'/g, "''")}'`,
    `-PidPath '${pidPath.replace(/'/g, "''")}'`,
    `-StatusPath '${statusPath.replace(/'/g, "''")}'`
  ].join(" "));
  const result = JSON.parse(fs.readFileSync(resultPath, "utf8").replace(/^\uFEFF/, ""));
  assert(result.ok === true, "Word COM 回归未成功");
  assert(fs.existsSync(result.outputDocx), "未生成 DOCX");
  assert(fs.existsSync(result.outputPdf), "未生成 PDF");
});

(async () => {
  let failures = 0;
  for (const item of tests) {
    try {
      await Promise.resolve(item.fn());
      console.log(`PASS ${item.name}`);
    } catch (error) {
      failures += 1;
      console.error(`FAIL ${item.name}`);
      console.error(error.message);
    }
  }

  if (failures) {
    console.error(`${failures} test(s) failed`);
    process.exit(1);
  }

  console.log(`${tests.length} test(s) passed`);
})();
