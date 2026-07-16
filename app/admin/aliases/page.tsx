'use client';

import { useEffect, useMemo, useState } from 'react';
import Image from 'next/image';
import Link from 'next/link';
import type { AliasEntry } from '@/lib/aliases';

interface DraftEntry extends AliasEntry {
  /** Local-only flag — true while the row hasn't been saved at least once. */
  isNew?: boolean;
  /** Local-only field — true while a save is in flight. */
  saving?: boolean;
  /** Last save error message, cleared on next edit. */
  error?: string;
}

function isValidEmail(s: string): boolean {
  if (!s) return false;
  return /\S+@\S+\.\S+/.test(s.trim());
}

export default function AliasesPage() {
  const [entries, setEntries] = useState<DraftEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [search, setSearch] = useState('');
  const [filter, setFilter] = useState<'all' | 'missing' | 'set'>('all');

  // ── Load on mount ─────────────────────────────────────────────────────────
  const reload = async () => {
    setLoading(true);
    setLoadError('');
    try {
      const res = await fetch('/api/aliases', { cache: 'no-store' });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Failed to load aliases');
      const list = (data.entries ?? []) as AliasEntry[];
      setEntries(list.map((e) => ({ ...e })));
    } catch (err) {
      setLoadError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    reload();
  }, []);

  // ── Helpers ───────────────────────────────────────────────────────────────

  const updateRow = (dataEmail: string, patch: Partial<DraftEntry>) => {
    setEntries((prev) =>
      prev.map((e) => (e.dataEmail === dataEmail ? { ...e, ...patch, error: undefined } : e))
    );
  };

  const saveRow = async (dataEmail: string) => {
    const row = entries.find((e) => e.dataEmail === dataEmail);
    if (!row) return;
    if (row.aliasEmail && !isValidEmail(row.aliasEmail)) {
      updateRow(dataEmail, { error: 'Alias email looks invalid' });
      return;
    }
    updateRow(dataEmail, { saving: true });
    try {
      const res = await fetch('/api/aliases', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify({
          dataEmail: row.dataEmail,
          aliasEmail: row.aliasEmail ?? '',
          displayName: row.displayName,
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Save failed');
      const saved = data.saved as AliasEntry;
      updateRow(dataEmail, { ...saved, saving: false, isNew: false });
    } catch (err) {
      updateRow(dataEmail, {
        saving: false,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  };

  const deleteRow = async (dataEmail: string) => {
    if (!window.confirm(`Remove alias entry for "${dataEmail}"?\n\nThe rep will still be in your data, but no alias will be applied.`)) {
      return;
    }
    updateRow(dataEmail, { saving: true });
    try {
      const res = await fetch(`/api/aliases?dataEmail=${encodeURIComponent(dataEmail)}`, {
        method: 'DELETE',
        cache: 'no-store',
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error ?? `Delete failed (${res.status})`);
      }
      setEntries((prev) => prev.filter((e) => e.dataEmail !== dataEmail));
    } catch (err) {
      updateRow(dataEmail, {
        saving: false,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  };

  const addBlankRow = () => {
    // If there's already an unsaved blank row, scroll to it instead of adding another
    const existingBlank = entries.find((e) => e.isNew && !e.dataEmail);
    if (existingBlank) {
      const el = document.getElementById('row-new-blank');
      el?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      return;
    }
    const blank: DraftEntry = {
      dataEmail: '',
      aliasEmail: '',
      displayName: '',
      addedAt: new Date().toISOString(),
      isNew: true,
    };
    setEntries((prev) => [blank, ...prev]);
  };

  const saveBlankRow = async (tmpRow: DraftEntry, dataEmail: string, aliasEmail: string, displayName: string) => {
    if (!isValidEmail(dataEmail)) {
      setEntries((prev) =>
        prev.map((e) => (e === tmpRow ? { ...e, error: 'Data email is required and must be valid' } : e))
      );
      return;
    }
    if (aliasEmail && !isValidEmail(aliasEmail)) {
      setEntries((prev) =>
        prev.map((e) => (e === tmpRow ? { ...e, error: 'Alias email looks invalid' } : e))
      );
      return;
    }
    setEntries((prev) =>
      prev.map((e) => (e === tmpRow ? { ...e, saving: true, error: undefined } : e))
    );
    try {
      const res = await fetch('/api/aliases', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify({ dataEmail, aliasEmail, displayName }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Save failed');
      const saved = data.saved as AliasEntry;
      // Replace the blank row with the saved one (and dedupe if it already existed)
      setEntries((prev) => {
        const withoutBlank = prev.filter((e) => e !== tmpRow);
        const withoutDup = withoutBlank.filter((e) => e.dataEmail !== saved.dataEmail);
        return [{ ...saved }, ...withoutDup];
      });
    } catch (err) {
      setEntries((prev) =>
        prev.map((e) =>
          e === tmpRow
            ? { ...e, saving: false, error: err instanceof Error ? err.message : String(err) }
            : e
        )
      );
    }
  };

  // ── Filtered view ─────────────────────────────────────────────────────────

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase();
    return entries.filter((e) => {
      if (filter === 'missing' && (e.aliasEmail ?? '').trim()) return false;
      if (filter === 'set' && !(e.aliasEmail ?? '').trim()) return false;
      if (!q) return true;
      return (
        e.dataEmail.toLowerCase().includes(q) ||
        (e.aliasEmail ?? '').toLowerCase().includes(q) ||
        (e.displayName ?? '').toLowerCase().includes(q)
      );
    });
  }, [entries, search, filter]);

  const stats = useMemo(() => {
    const total = entries.filter((e) => !e.isNew).length;
    const withAlias = entries.filter((e) => !e.isNew && (e.aliasEmail ?? '').trim()).length;
    return { total, withAlias, missing: total - withAlias };
  }, [entries]);

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen flex flex-col bg-gray-50">
      {/* Header */}
      <header className="bg-white shadow-sm border-b-4 border-[#DA291C]">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <Image src="/vital-logo.png" alt="Vital" width={120} height={48} className="object-contain" />
          <div className="flex items-center gap-6">
            <Link href="/" className="text-sm text-[#32373C] hover:text-[#DA291C] font-medium transition-colors">
              ← Score Card Builder
            </Link>
            <Link href="/admin/photo-map" className="text-sm text-[#32373C] hover:text-[#DA291C] font-medium transition-colors">
              Photo Mapping
            </Link>
            <Link href="/log" className="text-sm text-[#32373C] hover:text-[#DA291C] font-medium transition-colors">
              Activity Log
            </Link>
            <Image src="/perigee-logo.jpg" alt="Perigee" width={80} height={48} className="object-contain" />
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 max-w-6xl mx-auto w-full px-6 py-10">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-[#DA291C] mb-1">Email Aliases</h1>
          <p className="text-sm text-gray-500 max-w-3xl">
            Map each rep&apos;s data email (the address that appears in the Perigee export) to an alias —
            their current/active email. When &ldquo;Email to Reps&rdquo; is enabled, the rep&apos;s report is sent to
            <strong> both</strong> the data email and the alias.
          </p>
          <p className="text-xs text-gray-400 mt-2">
            New emails detected in uploaded data are added here automatically with no alias set.
          </p>
        </div>

        {/* Stats + actions */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-6 flex flex-wrap items-center gap-4">
          <div className="flex items-center gap-6 text-sm">
            <div>
              <span className="text-xs uppercase font-semibold text-gray-400 tracking-wide">Total</span>
              <div className="text-xl font-bold text-[#32373C]">{stats.total}</div>
            </div>
            <div>
              <span className="text-xs uppercase font-semibold text-gray-400 tracking-wide">With alias</span>
              <div className="text-xl font-bold text-green-600">{stats.withAlias}</div>
            </div>
            <div>
              <span className="text-xs uppercase font-semibold text-gray-400 tracking-wide">Missing</span>
              <div className="text-xl font-bold text-amber-600">{stats.missing}</div>
            </div>
          </div>

          <div className="flex-1" />

          <input
            type="search"
            placeholder="Search…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C] w-56"
          />

          <div className="inline-flex rounded-lg border border-gray-300 overflow-hidden text-xs">
            {(['all', 'missing', 'set'] as const).map((f) => (
              <button
                key={f}
                type="button"
                onClick={() => setFilter(f)}
                className={`px-3 py-2 font-semibold transition-colors ${
                  filter === f ? 'bg-[#DA291C] text-white' : 'bg-white text-gray-600 hover:bg-gray-50'
                }`}
              >
                {f === 'all' ? 'All' : f === 'missing' ? 'Missing only' : 'With alias'}
              </button>
            ))}
          </div>

          <button
            type="button"
            onClick={addBlankRow}
            className="px-4 py-2 bg-[#DA291C] hover:bg-[#B71C1C] text-white text-sm font-bold rounded-lg transition-colors flex items-center gap-1"
          >
            <span className="text-base leading-none">+</span> Add Entry
          </button>
        </div>

        {loading && (
          <div className="flex items-center gap-3 text-gray-500 py-10 justify-center">
            <svg className="animate-spin w-6 h-6 text-[#DA291C]" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
            </svg>
            Loading aliases…
          </div>
        )}

        {loadError && (
          <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-sm text-red-700">
            <strong>Error loading aliases:</strong> {loadError}
          </div>
        )}

        {!loading && !loadError && entries.length === 0 && (
          <div className="text-center py-16 text-gray-400 bg-white rounded-xl border border-gray-200">
            <p className="text-lg font-medium mb-1">No aliases yet</p>
            <p className="text-sm">Upload a Perigee export to populate this list automatically, or click <strong>+ Add Entry</strong> to add one manually.</p>
          </div>
        )}

        {!loading && !loadError && entries.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
            <table className="w-full text-left">
              <thead>
                <tr className="border-b border-gray-200 bg-[#32373C]">
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide w-[22%]">Display Name</th>
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide w-[30%]">Data Email</th>
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide w-[30%]">Alias Email</th>
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide text-right">Actions</th>
                </tr>
              </thead>
              <tbody>
                {visible.length === 0 && (
                  <tr>
                    <td colSpan={4} className="px-4 py-8 text-center text-gray-400 text-sm">
                      No entries match your filter.
                    </td>
                  </tr>
                )}
                {visible.map((row) =>
                  row.isNew && !row.dataEmail ? (
                    <BlankRow
                      key="new-blank"
                      row={row}
                      onSave={(de, ae, dn) => saveBlankRow(row, de, ae, dn)}
                      onCancel={() => setEntries((prev) => prev.filter((e) => e !== row))}
                    />
                  ) : (
                    <ExistingRow
                      key={row.dataEmail}
                      row={row}
                      onChange={(patch) => updateRow(row.dataEmail, patch)}
                      onSave={() => saveRow(row.dataEmail)}
                      onDelete={() => deleteRow(row.dataEmail)}
                    />
                  )
                )}
              </tbody>
            </table>
            <div className="px-4 py-2 border-t border-gray-100 bg-gray-50 text-xs text-gray-400 text-right">
              {visible.length} of {entries.length} entr{entries.length !== 1 ? 'ies' : 'y'}
            </div>
          </div>
        )}
      </main>

      <footer className="bg-white border-t border-gray-200 py-4 text-center text-xs text-gray-400">
        If you have any suggestions about how we might improve this tool, please email{' '}
        <a href="mailto:info@outerjoin.co.za" className="text-[#E8572A] underline">info@outerjoin.co.za</a>
      </footer>
    </div>
  );
}

// ── Row components ──────────────────────────────────────────────────────────

function ExistingRow({
  row,
  onChange,
  onSave,
  onDelete,
}: {
  row: DraftEntry;
  onChange: (patch: Partial<DraftEntry>) => void;
  onSave: () => void;
  onDelete: () => void;
}) {
  const hasAlias = !!(row.aliasEmail ?? '').trim();
  return (
    <tr className="border-b border-gray-100 hover:bg-gray-50 transition-colors align-top">
      <td className="px-4 py-3">
        <input
          type="text"
          value={row.displayName ?? ''}
          onChange={(e) => onChange({ displayName: e.target.value })}
          placeholder="(no name)"
          className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded focus:outline-none focus:border-[#DA291C]"
        />
      </td>
      <td className="px-4 py-3">
        <span className="text-sm text-gray-500 break-all">{row.dataEmail}</span>
      </td>
      <td className="px-4 py-3">
        <input
          type="email"
          value={row.aliasEmail ?? ''}
          onChange={(e) => onChange({ aliasEmail: e.target.value })}
          placeholder="new-email@example.com"
          className={`w-full px-2 py-1.5 text-sm border rounded focus:outline-none focus:border-[#DA291C] ${
            hasAlias ? 'border-green-300 bg-green-50/40' : 'border-amber-300 bg-amber-50/40'
          }`}
        />
        {row.error && <p className="mt-1 text-xs text-red-600">{row.error}</p>}
      </td>
      <td className="px-4 py-3 text-right whitespace-nowrap">
        <button
          type="button"
          onClick={onSave}
          disabled={row.saving}
          className="px-3 py-1.5 bg-[#DA291C] hover:bg-[#B71C1C] disabled:opacity-50 text-white text-xs font-bold rounded transition-colors mr-2"
        >
          {row.saving ? 'Saving…' : 'Save'}
        </button>
        <button
          type="button"
          onClick={onDelete}
          disabled={row.saving}
          className="px-3 py-1.5 bg-white border border-gray-300 hover:bg-gray-100 disabled:opacity-50 text-gray-600 text-xs font-bold rounded transition-colors"
        >
          Delete
        </button>
      </td>
    </tr>
  );
}

function BlankRow({
  row,
  onSave,
  onCancel,
}: {
  row: DraftEntry;
  onSave: (dataEmail: string, aliasEmail: string, displayName: string) => void;
  onCancel: () => void;
}) {
  const [dataEmail, setDataEmail] = useState('');
  const [aliasEmail, setAliasEmail] = useState('');
  const [displayName, setDisplayName] = useState('');
  return (
    <tr id="row-new-blank" className="border-b border-gray-100 bg-amber-50/40 align-top">
      <td className="px-4 py-3">
        <input
          type="text"
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          placeholder="Rep full name"
          className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded focus:outline-none focus:border-[#DA291C]"
        />
      </td>
      <td className="px-4 py-3">
        <input
          type="email"
          value={dataEmail}
          onChange={(e) => setDataEmail(e.target.value)}
          placeholder="email-as-it-appears-in-data@example.com"
          className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded focus:outline-none focus:border-[#DA291C]"
          autoFocus
        />
      </td>
      <td className="px-4 py-3">
        <input
          type="email"
          value={aliasEmail}
          onChange={(e) => setAliasEmail(e.target.value)}
          placeholder="new-or-alias-email@example.com (optional)"
          className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded focus:outline-none focus:border-[#DA291C]"
        />
        {row.error && <p className="mt-1 text-xs text-red-600">{row.error}</p>}
      </td>
      <td className="px-4 py-3 text-right whitespace-nowrap">
        <button
          type="button"
          onClick={() => onSave(dataEmail, aliasEmail, displayName)}
          disabled={row.saving}
          className="px-3 py-1.5 bg-[#DA291C] hover:bg-[#B71C1C] disabled:opacity-50 text-white text-xs font-bold rounded transition-colors mr-2"
        >
          {row.saving ? 'Saving…' : 'Add'}
        </button>
        <button
          type="button"
          onClick={onCancel}
          disabled={row.saving}
          className="px-3 py-1.5 bg-white border border-gray-300 hover:bg-gray-100 disabled:opacity-50 text-gray-600 text-xs font-bold rounded transition-colors"
        >
          Cancel
        </button>
      </td>
    </tr>
  );
}
