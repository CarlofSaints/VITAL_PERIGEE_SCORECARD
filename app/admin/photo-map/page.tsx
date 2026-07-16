'use client';

import { useEffect, useMemo, useState } from 'react';
import Image from 'next/image';
import Link from 'next/link';

interface MapEntry {
  photoQuestion: string;
  sheetName: string;
  addedAt: string;
  updatedAt?: string;
  isDefault?: boolean;
}

interface DraftEntry extends MapEntry {
  saving?: boolean;
  error?: string;
}

export default function PhotoMapPage() {
  const [entries, setEntries] = useState<DraftEntry[]>([]);
  const [sheets, setSheets] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState('');
  const [search, setSearch] = useState('');

  // New-entry form
  const [newValue, setNewValue] = useState('');
  const [newSheet, setNewSheet] = useState('');
  const [newSaving, setNewSaving] = useState(false);
  const [newError, setNewError] = useState('');

  const reload = async () => {
    setLoading(true);
    setLoadError('');
    try {
      const res = await fetch('/api/photo-map', { cache: 'no-store' });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Failed to load mappings');
      setEntries(((data.entries ?? []) as MapEntry[]).map((e) => ({ ...e })));
      setSheets((data.sheets ?? []) as string[]);
    } catch (err) {
      setLoadError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    reload();
  }, []);

  const updateRow = (photoQuestion: string, patch: Partial<DraftEntry>) => {
    setEntries((prev) =>
      prev.map((e) => (e.photoQuestion === photoQuestion ? { ...e, ...patch, error: undefined } : e))
    );
  };

  const saveRow = async (photoQuestion: string, sheetName: string) => {
    updateRow(photoQuestion, { saving: true, sheetName });
    try {
      const res = await fetch('/api/photo-map', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify({ photoQuestion, sheetName }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Save failed');
      const saved = data.saved as MapEntry;
      updateRow(photoQuestion, { ...saved, saving: false, isDefault: false });
    } catch (err) {
      updateRow(photoQuestion, { saving: false, error: err instanceof Error ? err.message : String(err) });
    }
  };

  const deleteRow = async (photoQuestion: string) => {
    if (!window.confirm(`Remove the mapping for "${photoQuestion}"?\n\nIt will be prompted for again next time this value appears in an upload.`)) {
      return;
    }
    updateRow(photoQuestion, { saving: true });
    try {
      const res = await fetch(`/api/photo-map?photoQuestion=${encodeURIComponent(photoQuestion)}`, {
        method: 'DELETE',
        cache: 'no-store',
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error ?? `Delete failed (${res.status})`);
      }
      await reload();
    } catch (err) {
      updateRow(photoQuestion, { saving: false, error: err instanceof Error ? err.message : String(err) });
    }
  };

  const addEntry = async () => {
    const value = newValue.trim();
    if (!value) { setNewError('Photo Question value is required'); return; }
    if (!newSheet) { setNewError('Pick a target sheet'); return; }
    setNewSaving(true);
    setNewError('');
    try {
      const res = await fetch('/api/photo-map', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        cache: 'no-store',
        body: JSON.stringify({ photoQuestion: value, sheetName: newSheet }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Save failed');
      setNewValue('');
      setNewSheet('');
      await reload();
    } catch (err) {
      setNewError(err instanceof Error ? err.message : String(err));
    } finally {
      setNewSaving(false);
    }
  };

  const visible = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return entries;
    return entries.filter(
      (e) => e.photoQuestion.toLowerCase().includes(q) || e.sheetName.toLowerCase().includes(q)
    );
  }, [entries, search]);

  const stats = useMemo(() => {
    const total = entries.length;
    const custom = entries.filter((e) => !e.isDefault).length;
    return { total, custom, defaults: total - custom };
  }, [entries]);

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
            <Link href="/admin/aliases" className="text-sm text-[#32373C] hover:text-[#DA291C] font-medium transition-colors">
              Email Aliases
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
          <h1 className="text-2xl font-bold text-[#DA291C] mb-1">Photo Question Mapping</h1>
          <p className="text-sm text-gray-500 max-w-3xl">
            Some Perigee photos have a <strong>Photo Question</strong> that isn&apos;t tied to a scorecard question
            (e.g. &ldquo;Photo stock pressure&rdquo;). On the Exception Report these photos are added to the sheet
            mapped here — matched to the store by Store Code, or added as a new line if the store isn&apos;t on that sheet.
          </p>
          <p className="text-xs text-gray-400 mt-2">
            Four values are mapped by default. New values you map while uploading are saved here automatically.
          </p>
        </div>

        {/* Stats */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-6 flex flex-wrap items-center gap-6">
          <div>
            <span className="text-xs uppercase font-semibold text-gray-400 tracking-wide">Total</span>
            <div className="text-xl font-bold text-[#32373C]">{stats.total}</div>
          </div>
          <div>
            <span className="text-xs uppercase font-semibold text-gray-400 tracking-wide">Custom</span>
            <div className="text-xl font-bold text-green-600">{stats.custom}</div>
          </div>
          <div>
            <span className="text-xs uppercase font-semibold text-gray-400 tracking-wide">Defaults</span>
            <div className="text-xl font-bold text-[#32373C]">{stats.defaults}</div>
          </div>
          <div className="flex-1" />
          <input
            type="search"
            placeholder="Search…"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C] w-56"
          />
        </div>

        {/* Add new mapping */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-4 mb-6">
          <h2 className="text-sm font-bold text-[#32373C] mb-3">Add a mapping</h2>
          <div className="flex flex-col sm:flex-row gap-3">
            <input
              type="text"
              value={newValue}
              onChange={(e) => { setNewValue(e.target.value); setNewError(''); }}
              placeholder='Photo Question value (e.g. "Photo stock pressure")'
              className="flex-1 px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C]"
            />
            <select
              value={newSheet}
              onChange={(e) => { setNewSheet(e.target.value); setNewError(''); }}
              className="w-full sm:w-64 px-3 py-2 border border-gray-300 rounded-lg text-sm bg-white focus:outline-none focus:border-[#DA291C]"
            >
              <option value="">— target sheet —</option>
              {sheets.map((s) => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
            <button
              type="button"
              onClick={addEntry}
              disabled={newSaving}
              className="px-5 py-2 bg-[#DA291C] hover:bg-[#B71C1C] disabled:opacity-50 text-white text-sm font-bold rounded-lg transition-colors whitespace-nowrap"
            >
              {newSaving ? 'Saving…' : 'Add'}
            </button>
          </div>
          {newError && <p className="mt-2 text-xs text-red-600">{newError}</p>}
        </div>

        {loading && (
          <div className="flex items-center gap-3 text-gray-500 py-10 justify-center">
            <svg className="animate-spin w-6 h-6 text-[#DA291C]" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
            </svg>
            Loading mappings…
          </div>
        )}

        {loadError && (
          <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-sm text-red-700">
            <strong>Error loading mappings:</strong> {loadError}
          </div>
        )}

        {!loading && !loadError && entries.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
            <table className="w-full text-left">
              <thead>
                <tr className="border-b border-gray-200 bg-[#32373C]">
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide w-[45%]">Photo Question</th>
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide w-[35%]">Target Sheet</th>
                  <th className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide text-right">Actions</th>
                </tr>
              </thead>
              <tbody>
                {visible.length === 0 && (
                  <tr>
                    <td colSpan={3} className="px-4 py-8 text-center text-gray-400 text-sm">
                      No mappings match your search.
                    </td>
                  </tr>
                )}
                {visible.map((row) => (
                  <tr key={row.photoQuestion} className="border-b border-gray-100 hover:bg-gray-50 transition-colors align-top">
                    <td className="px-4 py-3">
                      <span className="text-sm text-[#32373C] break-all">{row.photoQuestion}</span>
                      {row.isDefault && (
                        <span className="ml-2 text-[10px] uppercase tracking-wide font-semibold text-gray-400 border border-gray-200 rounded px-1.5 py-0.5">default</span>
                      )}
                    </td>
                    <td className="px-4 py-3">
                      <select
                        value={row.sheetName}
                        onChange={(e) => saveRow(row.photoQuestion, e.target.value)}
                        disabled={row.saving}
                        className="w-full px-2 py-1.5 text-sm border border-gray-300 rounded bg-white focus:outline-none focus:border-[#DA291C] disabled:opacity-50"
                      >
                        {sheets.map((s) => (
                          <option key={s} value={s}>{s}</option>
                        ))}
                      </select>
                      {row.error && <p className="mt-1 text-xs text-red-600">{row.error}</p>}
                    </td>
                    <td className="px-4 py-3 text-right whitespace-nowrap">
                      {row.isDefault ? (
                        <span className="text-xs text-gray-400 italic">built-in</span>
                      ) : (
                        <button
                          type="button"
                          onClick={() => deleteRow(row.photoQuestion)}
                          disabled={row.saving}
                          className="px-3 py-1.5 bg-white border border-gray-300 hover:bg-gray-100 disabled:opacity-50 text-gray-600 text-xs font-bold rounded transition-colors"
                        >
                          {row.saving ? '…' : 'Delete'}
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="px-4 py-2 border-t border-gray-100 bg-gray-50 text-xs text-gray-400 text-right">
              {visible.length} of {entries.length} mapping{entries.length !== 1 ? 's' : ''}
            </div>
          </div>
        )}

        {!loading && !loadError && entries.length === 0 && (
          <div className="text-center py-16 text-gray-400 bg-white rounded-xl border border-gray-200">
            <p className="text-lg font-medium mb-1">No mappings yet</p>
            <p className="text-sm">Add one above, or upload a Perigee export and map new Photo Question values as they appear.</p>
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
