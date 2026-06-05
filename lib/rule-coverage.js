const fs = require("fs");
const path = require("path");

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

function businessRuleLeaves(rules) {
  const metadata = new Set(["ruleId", "revisionId", "updatedAt", "note", "restoredFrom"]);
  return walkLeaves(rules).filter((item) => !metadata.has(item));
}

function patternToRegExp(pattern) {
  const escaped = pattern
    .split("*")
    .map((part) => part.replace(/[.+?^${}()|[\]\\]/g, "\\$&"))
    .join(".*");
  return new RegExp(`^${escaped}$`);
}

function matchPattern(pattern, leaves) {
  const regex = patternToRegExp(pattern);
  return leaves.filter((leaf) => regex.test(leaf));
}

function expandCoverageEntries(coverage, leaves) {
  const byPath = new Map();
  const unmatchedPatterns = [];

  for (const entry of coverage.entries || []) {
    for (const pattern of entry.paths || []) {
      const matches = matchPattern(pattern, leaves);
      if (!matches.length) {
        unmatchedPatterns.push({ entry: entry.id, pattern });
      }
      for (const leaf of matches) {
        if (!byPath.has(leaf)) byPath.set(leaf, []);
        byPath.get(leaf).push(entry);
      }
    }
  }

  return { byPath, unmatchedPatterns };
}

function loadJson(root, relativePath) {
  return JSON.parse(fs.readFileSync(path.join(root, relativePath), "utf8"));
}

module.exports = {
  businessRuleLeaves,
  expandCoverageEntries,
  loadJson,
  matchPattern,
  patternToRegExp,
  walkLeaves
};
