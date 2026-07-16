/**
 * Vital Score Card — Photo Question → Sheet Mapping Store
 *
 * The Perigee export contains "Photo …" columns that don't belong to any
 * scorecard question (e.g. "Photo stock pressure", "Photo POS"). On the
 * Exception Report these photos are distributed into the relevant per-question
 * sheet based on their Photo Question value.
 *
 * A small set of these values is known up-front (see DEFAULT_PHOTO_QUESTION_MAP)
 * and never needs mapping. Any NEW Photo Question value is surfaced in the UI so
 * the user can pick a target sheet; that choice is persisted here (a single JSON
 * file on SharePoint, alongside VITAL_EMAIL_ALIASES.json) so it's only asked once.
 *
 * Mirrors the persistence pattern in lib/aliases.ts.
 */

import { downloadPhotoMapFile, uploadPhotoMapFile } from '@/lib/sharepoint';
import {
  DEFAULT_PHOTO_QUESTION_MAP,
  SELECTABLE_SHEETS,
  normPhotoQuestion,
} from '@/constants/sheets';

export interface PhotoMapEntry {
  /** Photo Question value as it appears in the data (display form). */
  photoQuestion: string;
  /** Target exception-report sheet name this value's photos flow into. */
  sheetName: string;
  /** First time this mapping was recorded. */
  addedAt: string;
  /** Last time the target sheet was changed. */
  updatedAt?: string;
}

interface PhotoMapFile {
  version: 1;
  entries: PhotoMapEntry[];
}

const EMPTY_FILE: PhotoMapFile = { version: 1, entries: [] };

// In-memory cache so writes within the same request/instance are immediately
// visible to subsequent reads.
let _cache: PhotoMapFile | null = null;

function nowIso(): string {
  return new Date().toISOString();
}

// ── Load / save ───────────────────────────────────────────────────────────────

export async function loadPhotoMap(): Promise<PhotoMapFile> {
  if (_cache) return _cache;
  try {
    const buf = await downloadPhotoMapFile();
    if (!buf) {
      _cache = { ...EMPTY_FILE, entries: [] };
      return _cache;
    }
    const parsed = JSON.parse(buf.toString('utf8')) as PhotoMapFile;
    if (!parsed || !Array.isArray(parsed.entries)) {
      _cache = { ...EMPTY_FILE, entries: [] };
      return _cache;
    }
    parsed.entries = parsed.entries
      .filter((e) => e && e.photoQuestion && e.sheetName)
      .map((e) => ({ ...e, photoQuestion: e.photoQuestion.trim(), sheetName: e.sheetName.trim() }));
    _cache = parsed;
    return _cache;
  } catch (err) {
    console.error('[photoMap] loadPhotoMap failed:', err instanceof Error ? err.message : err);
    // Fall back to empty so the app keeps working (defaults still apply) even if SP is down.
    _cache = { ...EMPTY_FILE, entries: [] };
    return _cache;
  }
}

export async function savePhotoMap(file: PhotoMapFile): Promise<void> {
  const toWrite: PhotoMapFile = {
    version: 1,
    entries: file.entries
      .filter((e) => e.photoQuestion && e.sheetName)
      .map((e) => ({ ...e, photoQuestion: e.photoQuestion.trim(), sheetName: e.sheetName.trim() })),
  };
  _cache = toWrite;
  await uploadPhotoMapFile(Buffer.from(JSON.stringify(toWrite, null, 2), 'utf8'));
}

// ── Resolution ─────────────────────────────────────────────────────────────────

/**
 * The effective mapping: built-in defaults overlaid with the persisted store
 * (store wins). Keyed by normalised Photo Question → sheet name.
 */
export async function resolvePhotoMap(): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  for (const [key, sheet] of Object.entries(DEFAULT_PHOTO_QUESTION_MAP)) {
    map.set(key, sheet);
  }
  const { entries } = await loadPhotoMap();
  for (const e of entries) {
    map.set(normPhotoQuestion(e.photoQuestion), e.sheetName);
  }
  return map;
}

/**
 * For a list of Photo Question values found in an upload, resolve each to its
 * target sheet (or null if unknown). De-duplicated, original display form kept.
 */
export async function resolvePhotoQuestions(
  values: string[]
): Promise<{ value: string; sheet: string | null }[]> {
  const map = await resolvePhotoMap();
  const seen = new Set<string>();
  const out: { value: string; sheet: string | null }[] = [];
  for (const v of values) {
    const key = normPhotoQuestion(v);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push({ value: v.trim(), sheet: map.get(key) ?? null });
  }
  return out;
}

// ── Mutations ────────────────────────────────────────────────────────────────

/** All entries as shown in the admin UI: defaults (read-only) + persisted. */
export async function listPhotoMapEntries(): Promise<
  (PhotoMapEntry & { isDefault?: boolean })[]
> {
  const { entries } = await loadPhotoMap();
  const storedKeys = new Set(entries.map((e) => normPhotoQuestion(e.photoQuestion)));
  const defaults: (PhotoMapEntry & { isDefault?: boolean })[] = Object.entries(
    DEFAULT_PHOTO_QUESTION_MAP
  )
    // Only show a default when the user hasn't overridden it in the store.
    .filter(([key]) => !storedKeys.has(key))
    .map(([key, sheet]) => ({
      photoQuestion: key,
      sheetName: sheet,
      addedAt: '',
      isDefault: true,
    }));
  return [...defaults, ...entries];
}

/** Add or update a Photo Question → sheet mapping. Returns the saved entry. */
export async function upsertPhotoMapping(input: {
  photoQuestion: string;
  sheetName: string;
}): Promise<PhotoMapEntry> {
  const photoQuestion = (input.photoQuestion ?? '').trim();
  const sheetName = (input.sheetName ?? '').trim();
  if (!photoQuestion) throw new Error('photoQuestion is required');
  if (!sheetName) throw new Error('sheetName is required');
  if (!SELECTABLE_SHEETS.includes(sheetName)) {
    throw new Error(`Unknown sheet "${sheetName}"`);
  }

  const file = await loadPhotoMap();
  const key = normPhotoQuestion(photoQuestion);
  const idx = file.entries.findIndex((e) => normPhotoQuestion(e.photoQuestion) === key);
  if (idx >= 0) {
    const updated: PhotoMapEntry = {
      ...file.entries[idx],
      photoQuestion,
      sheetName,
      updatedAt: nowIso(),
    };
    file.entries[idx] = updated;
    await savePhotoMap(file);
    return updated;
  }
  const created: PhotoMapEntry = { photoQuestion, sheetName, addedAt: nowIso() };
  file.entries.push(created);
  await savePhotoMap(file);
  return created;
}

/**
 * Persist a batch of mappings, skipping ones that already resolve to the same
 * target (so re-submitting known defaults doesn't churn the store).
 * Returns the entries that were actually written.
 */
export async function upsertPhotoMappings(
  mappings: { photoQuestion: string; sheetName: string }[]
): Promise<PhotoMapEntry[]> {
  const current = await resolvePhotoMap();
  const saved: PhotoMapEntry[] = [];
  for (const m of mappings) {
    const pq = (m.photoQuestion ?? '').trim();
    const sheet = (m.sheetName ?? '').trim();
    if (!pq || !sheet || !SELECTABLE_SHEETS.includes(sheet)) continue;
    if (current.get(normPhotoQuestion(pq)) === sheet) continue; // already effective
    saved.push(await upsertPhotoMapping({ photoQuestion: pq, sheetName: sheet }));
  }
  return saved;
}

/** Remove a persisted mapping by Photo Question value. Defaults can't be removed. */
export async function removePhotoMapping(photoQuestion: string): Promise<boolean> {
  const file = await loadPhotoMap();
  const key = normPhotoQuestion(photoQuestion);
  const before = file.entries.length;
  file.entries = file.entries.filter((e) => normPhotoQuestion(e.photoQuestion) !== key);
  if (file.entries.length === before) return false;
  await savePhotoMap(file);
  return true;
}
