/**
 * Vital Score Card — Email Alias Store
 *
 * Maintains a single JSON file (VITAL_EMAIL_ALIASES.json) on SharePoint in the
 * Historical Data folder, mapping each "data email" (the address that appears
 * in the Perigee export for a rep) to one optional "alias email" — typically
 * the rep's current/active address.
 *
 * When a rep report is sent, BOTH addresses receive the email so we can keep
 * delivering during an email transition without losing anyone.
 *
 * Persistence pattern follows the same in-memory `_cache` approach used in
 * other Vercel-deployed projects so writes are immediately readable on the
 * same instance (Vercel env-var stale-read workaround in spirit, even though
 * here the canonical store is on SharePoint and not env vars).
 */

import { downloadAliasesFile, uploadAliasesFile } from '@/lib/sharepoint';

export interface AliasEntry {
  /** Email address as it appears in the Perigee data — case-insensitive key. */
  dataEmail: string;
  /** Optional second address that should also receive the rep's report. */
  aliasEmail: string;
  /** Display name (e.g. rep's full name). Used in admin UI only. */
  displayName?: string;
  /** First time we saw this data email (auto-discovered or manually added). */
  addedAt: string;
  /** Last time the alias email was updated. */
  updatedAt?: string;
}

interface AliasFile {
  version: 1;
  entries: AliasEntry[];
}

const EMPTY_FILE: AliasFile = { version: 1, entries: [] };

// In-memory cache so writes within the same request/instance are immediately
// visible to subsequent reads (no stale-read across handlers).
let _cache: AliasFile | null = null;

function normEmail(e: string): string {
  return (e ?? '').trim().toLowerCase();
}

function nowIso(): string {
  return new Date().toISOString();
}

// ── Load / save ───────────────────────────────────────────────────────────────

export async function loadAliases(): Promise<AliasFile> {
  if (_cache) return _cache;
  try {
    const buf = await downloadAliasesFile();
    if (!buf) {
      _cache = { ...EMPTY_FILE, entries: [] };
      return _cache;
    }
    const parsed = JSON.parse(buf.toString('utf8')) as AliasFile;
    if (!parsed || !Array.isArray(parsed.entries)) {
      _cache = { ...EMPTY_FILE, entries: [] };
      return _cache;
    }
    // Normalise emails on read to be safe
    parsed.entries = parsed.entries.map((e) => ({
      ...e,
      dataEmail:  normEmail(e.dataEmail),
      aliasEmail: (e.aliasEmail ?? '').trim(),
    }));
    _cache = parsed;
    return _cache;
  } catch (err) {
    console.error('[aliases] loadAliases failed:', err instanceof Error ? err.message : err);
    // Fall back to empty so the app keeps working even if SP is down
    _cache = { ...EMPTY_FILE, entries: [] };
    return _cache;
  }
}

export async function saveAliases(file: AliasFile): Promise<void> {
  const toWrite: AliasFile = {
    version: 1,
    entries: file.entries.map((e) => ({
      ...e,
      dataEmail:  normEmail(e.dataEmail),
      aliasEmail: (e.aliasEmail ?? '').trim(),
    })),
  };
  _cache = toWrite; // update cache immediately
  await uploadAliasesFile(Buffer.from(JSON.stringify(toWrite, null, 2), 'utf8'));
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Look up the alias email (if any) for a given data email. */
export async function getAliasFor(dataEmail: string): Promise<string | null> {
  const { entries } = await loadAliases();
  const key = normEmail(dataEmail);
  const hit = entries.find((e) => e.dataEmail === key);
  if (!hit) return null;
  const a = (hit.aliasEmail ?? '').trim();
  return a ? a : null;
}

/** Returns data emails that are NOT yet present in the alias store. */
export async function findUnknownEmails(
  emails: { dataEmail: string; displayName?: string }[]
): Promise<{ dataEmail: string; displayName?: string }[]> {
  const { entries } = await loadAliases();
  const known = new Set(entries.map((e) => e.dataEmail));
  const seen = new Set<string>();
  const out: { dataEmail: string; displayName?: string }[] = [];
  for (const e of emails) {
    const key = normEmail(e.dataEmail);
    if (!key || known.has(key) || seen.has(key)) continue;
    seen.add(key);
    out.push({ dataEmail: key, displayName: e.displayName });
  }
  return out;
}

/**
 * Adds new entries (without aliasEmail) for any emails not yet known.
 * Used by the parse route to auto-populate the store as data is uploaded.
 * Existing entries are left untouched. Returns the list of newly added entries.
 */
export async function autoRegisterEmails(
  emails: { dataEmail: string; displayName?: string }[]
): Promise<AliasEntry[]> {
  const file = await loadAliases();
  const known = new Set(file.entries.map((e) => e.dataEmail));
  const added: AliasEntry[] = [];
  for (const e of emails) {
    const key = normEmail(e.dataEmail);
    if (!key || known.has(key)) continue;
    const entry: AliasEntry = {
      dataEmail:  key,
      aliasEmail: '',
      displayName: e.displayName,
      addedAt: nowIso(),
    };
    file.entries.push(entry);
    added.push(entry);
    known.add(key);
  }
  if (added.length > 0) {
    await saveAliases(file);
  }
  return added;
}

/**
 * Add a new entry or update the alias / displayName of an existing one.
 * Returns the saved entry.
 */
export async function upsertAlias(input: {
  dataEmail: string;
  aliasEmail?: string;
  displayName?: string;
}): Promise<AliasEntry> {
  const file = await loadAliases();
  const key = normEmail(input.dataEmail);
  if (!key) throw new Error('dataEmail is required');

  const idx = file.entries.findIndex((e) => e.dataEmail === key);
  if (idx >= 0) {
    const current = file.entries[idx];
    const updated: AliasEntry = {
      ...current,
      aliasEmail:  (input.aliasEmail ?? current.aliasEmail ?? '').trim(),
      displayName: input.displayName ?? current.displayName,
      updatedAt:   nowIso(),
    };
    file.entries[idx] = updated;
    await saveAliases(file);
    return updated;
  }

  const created: AliasEntry = {
    dataEmail:   key,
    aliasEmail:  (input.aliasEmail ?? '').trim(),
    displayName: input.displayName,
    addedAt:     nowIso(),
  };
  file.entries.push(created);
  await saveAliases(file);
  return created;
}

/** Remove an entry by data email. Returns true if anything was removed. */
export async function removeAlias(dataEmail: string): Promise<boolean> {
  const file = await loadAliases();
  const key = normEmail(dataEmail);
  const before = file.entries.length;
  file.entries = file.entries.filter((e) => e.dataEmail !== key);
  if (file.entries.length === before) return false;
  await saveAliases(file);
  return true;
}
