const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");

const DEFAULT_RETENTION_DAYS = 30;
const DEFAULT_MAX_JOBS = 200;

function parseNonNegativeInteger(value, fallback) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) return fallback;
  return Math.floor(number);
}

function normalizeCleanupPolicy(input = {}, env = process.env) {
  return {
    retentionDays: parseNonNegativeInteger(input.retentionDays ?? env.REDHEAD_JOB_RETENTION_DAYS, DEFAULT_RETENTION_DAYS),
    maxJobs: parseNonNegativeInteger(input.maxJobs ?? env.REDHEAD_JOB_RETENTION_MAX, DEFAULT_MAX_JOBS),
    dryRun: Boolean(input.dryRun)
  };
}

function isPathInside(parent, child) {
  const relative = path.relative(path.resolve(parent), path.resolve(child));
  return relative && !relative.startsWith("..") && !path.isAbsolute(relative);
}

async function pathExists(targetPath) {
  try {
    await fsp.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function getDirectorySize(dirPath) {
  let total = 0;
  async function walk(currentPath) {
    const entries = await fsp.readdir(currentPath, { withFileTypes: true });
    for (const entry of entries) {
      const childPath = path.join(currentPath, entry.name);
      if (entry.isDirectory()) {
        await walk(childPath);
      } else if (entry.isFile()) {
        const stat = await fsp.stat(childPath);
        total += stat.size;
      }
    }
  }
  if (await pathExists(dirPath)) await walk(dirPath);
  return total;
}

async function listJobDirectories(baseDir) {
  if (!(await pathExists(baseDir))) return [];
  const entries = await fsp.readdir(baseDir, { withFileTypes: true });
  const jobs = [];
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const dirPath = path.join(baseDir, entry.name);
    const stat = await fsp.stat(dirPath);
    jobs.push({
      name: entry.name,
      path: dirPath,
      mtimeMs: stat.mtimeMs,
      updatedAt: stat.mtime.toISOString(),
      sizeBytes: await getDirectorySize(dirPath)
    });
  }
  return jobs.sort((a, b) => b.mtimeMs - a.mtimeMs);
}

function selectCleanupTargets(outputJobs, uploadJobs, policy, now = new Date(), protectedNames = new Set()) {
  const cutoffMs = policy.retentionDays > 0 ? now.getTime() - policy.retentionDays * 24 * 60 * 60 * 1000 : null;
  const uploadByName = new Map(uploadJobs.map((job) => [job.name, job]));
  const outputTargets = [];
  const uploadTargets = [];
  const selectedUploads = new Set();

  outputJobs.forEach((job, index) => {
    if (protectedNames.has(job.name)) return;
    const tooOld = cutoffMs !== null && job.mtimeMs < cutoffMs;
    const tooMany = policy.maxJobs > 0 && index >= policy.maxJobs;
    if (!tooOld && !tooMany) return;

    outputTargets.push({ ...job, reason: tooOld ? "retentionDays" : "maxJobs" });
    const uploadJob = uploadByName.get(job.name);
    if (uploadJob && !protectedNames.has(uploadJob.name)) {
      uploadTargets.push({ ...uploadJob, reason: "matchingOutput" });
      selectedUploads.add(uploadJob.name);
    }
  });

  for (const job of uploadJobs) {
    if (protectedNames.has(job.name) || selectedUploads.has(job.name)) continue;
    if (outputJobs.some((outputJob) => outputJob.name === job.name)) continue;
    const orphanTooOld = cutoffMs !== null && job.mtimeMs < cutoffMs;
    if (orphanTooOld) {
      uploadTargets.push({ ...job, reason: "orphanUpload" });
      selectedUploads.add(job.name);
    }
  }

  return { outputTargets, uploadTargets };
}

async function removeDirectory(baseDir, targetPath) {
  if (!isPathInside(baseDir, targetPath)) {
    throw new Error(`拒绝清理非目标目录：${targetPath}`);
  }
  await fsp.rm(targetPath, { recursive: true, force: true });
}

async function cleanupJobArtifacts(options) {
  const outputDir = options.outputDir;
  const uploadDir = options.uploadDir;
  const policy = normalizeCleanupPolicy(options.policy || {});
  const protectedNames = new Set(options.protectedNames || []);
  const now = options.now || new Date();
  const outputJobs = await listJobDirectories(outputDir);
  const uploadJobs = await listJobDirectories(uploadDir);
  const { outputTargets, uploadTargets } = selectCleanupTargets(outputJobs, uploadJobs, policy, now, protectedNames);
  const targets = { outputs: outputTargets, uploads: uploadTargets };
  const reclaimedBytes = [...outputTargets, ...uploadTargets].reduce((sum, item) => sum + item.sizeBytes, 0);

  if (!policy.dryRun) {
    for (const target of outputTargets) await removeDirectory(outputDir, target.path);
    for (const target of uploadTargets) await removeDirectory(uploadDir, target.path);
  }

  return {
    policy,
    dryRun: policy.dryRun,
    scanned: {
      outputs: outputJobs.length,
      uploads: uploadJobs.length
    },
    targets,
    deleted: policy.dryRun ? { outputs: [], uploads: [] } : targets,
    targetCounts: {
      outputs: outputTargets.length,
      uploads: uploadTargets.length
    },
    reclaimedBytes,
    protectedNames: [...protectedNames]
  };
}

module.exports = {
  cleanupJobArtifacts,
  getDirectorySize,
  listJobDirectories,
  normalizeCleanupPolicy,
  selectCleanupTargets
};
