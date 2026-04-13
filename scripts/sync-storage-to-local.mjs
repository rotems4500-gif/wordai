import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const projectRoot = process.cwd();
const targetDir = path.join(projectRoot, "public", "project-materials");
const targetIndexPath = path.join(targetDir, "index.json");

function toPublicDownloadUrl(rawUrl, fallbackObjectPath = "") {
  const trimmed = String(rawUrl || "").trim();
  if (!trimmed) return "";
  if (trimmed.startsWith("gs://")) {
    const noScheme = trimmed.slice(5);
    const slashAt = noScheme.indexOf("/");
    if (slashAt < 0) {
      if (!fallbackObjectPath) return "";
      const bucketOnly = noScheme.trim();
      if (!bucketOnly) return "";
      return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucketOnly)}/o/${encodeURIComponent(fallbackObjectPath)}?alt=media`;
    }
    if (slashAt === 0) return "";
    const bucket = noScheme.slice(0, slashAt);
    const objectPath = noScheme.slice(slashAt + 1) || fallbackObjectPath;
    if (!objectPath) return "";
    return `https://firebasestorage.googleapis.com/v0/b/${encodeURIComponent(bucket)}/o/${encodeURIComponent(objectPath)}?alt=media`;
  }
  return trimmed;
}

function inferFileNameFromUrl(url) {
  try {
    const parsed = new URL(url);
    const pathname = decodeURIComponent(parsed.pathname || "");
    const lastSegment = pathname.split("/").filter(Boolean).pop() || "material.bin";
    return lastSegment.includes(".") ? lastSegment : `${lastSegment}.bin`;
  } catch {
    return "material.bin";
  }
}

function sanitizeRelativePath(input) {
  const normalized = String(input || "").trim().replace(/\\/g, "/");
  const cleaned = normalized.replace(/^\/+/, "");
  if (!cleaned) return "";
  const segments = cleaned.split("/").filter(Boolean);
  if (segments.some((segment) => segment === "." || segment === "..")) {
    throw new Error(`Invalid path in index entry: ${input}`);
  }
  return segments.join("/");
}

function getBucketFromDownloadUrl(indexUrl) {
  try {
    const parsed = new URL(indexUrl);
    const match = parsed.pathname.match(/\/v0\/b\/([^/]+)\/o\//);
    return match ? decodeURIComponent(match[1]) : "";
  } catch {
    return "";
  }
}

async function fetchMaterialsFromIndexUrl(indexUrl) {
  const response = await fetch(indexUrl);
  if (!response.ok) {
    throw new Error(`Failed to fetch index (${response.status})`);
  }

  const rows = await response.json();
  if (!Array.isArray(rows)) {
    throw new Error("Storage index must be a JSON array");
  }

  const baseUrl = new URL(indexUrl);
  const storageBucket = getBucketFromDownloadUrl(indexUrl);
  return rows.map((row, idx) => {
    const filePath = String(row.file || row.path || "").trim();
    const directUrlRaw = String(row.url || filePath || "").trim();
    if (!directUrlRaw) {
      throw new Error(`Index row ${idx + 1} is missing file/path/url`);
    }

    let directUrl = "";
    if (directUrlRaw.startsWith("gs://")) {
      directUrl = toPublicDownloadUrl(directUrlRaw);
    } else if (row.url) {
      directUrl = new URL(directUrlRaw, baseUrl).toString();
    } else if (storageBucket && filePath) {
      directUrl = toPublicDownloadUrl(`gs://${storageBucket}/${filePath}`);
    } else {
      directUrl = new URL(directUrlRaw, baseUrl).toString();
    }

    const relativeFile = sanitizeRelativePath(filePath || inferFileNameFromUrl(directUrl));
    const fileName = path.posix.basename(relativeFile);
    return {
      id: String(row.id || relativeFile || `material-${idx + 1}`),
      title: String(row.title || fileName),
      type: String(row.type || path.posix.extname(fileName).slice(1) || ""),
      file: relativeFile,
      sourceUrl: directUrl,
      syncedFrom: String(row.syncedFrom || directUrl),
    };
  });
}

async function ensureDirForFile(filePath) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
}

async function downloadFile(url, destinationPath) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to download ${url} (${response.status})`);
  }
  const buffer = Buffer.from(await response.arrayBuffer());
  await ensureDirForFile(destinationPath);
  await fs.writeFile(destinationPath, buffer);
}

async function readExistingLocalIndex() {
  try {
    const raw = await fs.readFile(targetIndexPath, 'utf8');
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function mergeLocalIndexEntries(existingEntries, incomingEntries) {
  const merged = [];
  const seen = new Set();
  for (const entry of [...existingEntries, ...incomingEntries]) {
    const key = String(entry.file || entry.id || '').trim().toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    merged.push(entry);
  }
  return merged;
}

function buildSingleMaterialEntry(sourceUrl) {
  const fileName = inferFileNameFromUrl(sourceUrl);
  const relativeFile = sanitizeRelativePath(fileName);
  return [{
    id: relativeFile,
    title: fileName,
    type: path.posix.extname(fileName).slice(1) || "",
    file: relativeFile,
    sourceUrl,
    syncedFrom: sourceUrl,
  }];
}

function upsertLocalIndexEntries(existingEntries, incomingEntries) {
  const byFile = new Map();
  for (const entry of existingEntries) {
    const key = String(entry.file || entry.id || '').trim().toLowerCase();
    if (!key) continue;
    byFile.set(key, entry);
  }
  for (const entry of incomingEntries) {
    const key = String(entry.file || entry.id || '').trim().toLowerCase();
    if (!key) continue;
    byFile.set(key, entry);
  }
  return [...byFile.values()];
}

async function main() {
  const sourceArg = process.argv[2];
  const jsonMode = process.argv.includes('--json');
  if (!sourceArg) {
    console.error("Usage: node scripts/sync-storage-to-local.mjs <gs://bucket[/index.json] | https://.../index.json>");
    process.exit(1);
  }

  const sourceUrl = /^gs:\/\/[^/]+\/?$/i.test(sourceArg)
    ? toPublicDownloadUrl(sourceArg, "index.json")
    : toPublicDownloadUrl(sourceArg);

  if (!sourceUrl) {
    throw new Error("Invalid storage source URL");
  }

  const materials = /\.json(\?|$)/i.test(sourceUrl)
    ? await fetchMaterialsFromIndexUrl(sourceUrl)
    : buildSingleMaterialEntry(sourceUrl);
  if (!materials.length) {
    throw new Error("No materials found in remote index");
  }

  await fs.mkdir(targetDir, { recursive: true });

  for (const material of materials) {
    const destination = path.join(targetDir, material.file);
    if (!jsonMode) console.log(`Downloading ${material.title} -> ${material.file}`);
    await downloadFile(material.sourceUrl, destination);
  }

  const localIndex = materials.map(({ sourceUrl, ...rest }) => rest);
  const finalIndex = /\.json(\?|$)/i.test(sourceUrl)
    ? localIndex
    : upsertLocalIndexEntries(await readExistingLocalIndex(), localIndex);
  await fs.writeFile(targetIndexPath, JSON.stringify(finalIndex, null, 2) + "\n", "utf8");

  const result = {
    syncedFiles: materials.map((material) => material.file),
    count: materials.length,
    indexCount: finalIndex.length,
  };

  if (jsonMode) {
    process.stdout.write(JSON.stringify(result));
    return;
  }

  console.log(`Wrote ${finalIndex.length} entries to ${path.relative(projectRoot, targetIndexPath)}`);
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
