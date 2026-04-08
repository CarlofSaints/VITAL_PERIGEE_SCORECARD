'use client';

import { useState, useEffect, useRef, useCallback } from 'react';
import Image from 'next/image';
import Link from 'next/link';
import type { Rep } from '@/types';

// ── Comparison Report section ──────────────────────────────────────────────────

type CompPeriod = 'weekly' | 'monthly' | 'quarterly';
type CompAction = 'sharepoint' | 'email' | 'both';

interface CompResult {
  ok?: boolean;
  fileName?: string;
  spUrl?: string;
  spError?: string;
  emailSent?: boolean;
  emailError?: string;
  dateRange?: string;
  error?: string;
}

interface FilterOptions {
  channels:  string[];
  provinces: string[];
  reps:      { email: string; name: string }[];
  stores:    { name: string; channel: string }[];
}

function ComparisonReport() {
  const [from, setFrom]           = useState('');
  const [to, setTo]               = useState('');
  const [period, setPeriod]       = useState<CompPeriod>('monthly');
  const [action, setAction]       = useState<CompAction>('sharepoint');
  const [email, setEmail]         = useState('');
  const [emailConfirmed, setEmailConfirmed] = useState(false);
  const [running, setRunning]     = useState(false);
  const [result, setResult]       = useState<CompResult | null>(null);

  const needsEmail = action === 'email' || action === 'both';
  const emailValid = !needsEmail || (email.includes('@') && email.includes('.'));
  const canRun = !!(from && to && from <= to && emailValid && (!needsEmail || emailConfirmed));

  // Filter options loaded from history
  const [filterOptions, setFilterOptions]   = useState<FilterOptions | null>(null);
  const [filtersLoading, setFiltersLoading] = useState(true);
  const [selChannels, setSelChannels]       = useState<string[]>([]);
  const [selProvinces, setSelProvinces]     = useState<string[]>([]);
  const [selReps, setSelReps]               = useState<string[]>([]);
  const [selStores, setSelStores]           = useState<string[]>([]);

  // Load filter options from history on mount
  useEffect(() => {
    setFiltersLoading(true);
    fetch('/api/comparison')
      .then((r) => r.json())
      .then((data: FilterOptions & { error?: string }) => {
        if (!data.error) setFilterOptions(data);
      })
      .catch(() => {})
      .finally(() => setFiltersLoading(false));
  }, []);

  // Stores available based on selected channels (all stores when no channel filter)
  const availableStoreNames = filterOptions
    ? (selChannels.length > 0
        ? filterOptions.stores.filter((s) => selChannels.includes(s.channel)).map((s) => s.name)
        : filterOptions.stores.map((s) => s.name))
    : [];

  // When channel selection changes, deselect stores that are no longer available
  useEffect(() => {
    if (!filterOptions) return;
    const available = selChannels.length > 0
      ? filterOptions.stores.filter((s) => selChannels.includes(s.channel)).map((s) => s.name)
      : filterOptions.stores.map((s) => s.name);
    setSelStores((prev) => prev.filter((s) => available.includes(s)));
  }, [selChannels, filterOptions]);

  const handleRun = async () => {
    if (!canRun) return;
    setRunning(true);
    setResult(null);
    try {
      const res = await fetch('/api/comparison', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          from, to, period, action,
          emailAddress: needsEmail ? email : undefined,
          filters: {
            channels:  selChannels.length  ? selChannels  : undefined,
            provinces: selProvinces.length ? selProvinces : undefined,
            reps:      selReps.length      ? selReps      : undefined,
            stores:    selStores.length    ? selStores    : undefined,
          },
        }),
      });
      const data = await res.json();
      setResult(data);
    } catch (err: unknown) {
      setResult({ error: err instanceof Error ? err.message : 'Unknown error' });
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mt-8">
      <h2 className="text-base font-bold text-[#32373C] mb-1 flex items-center gap-2">
        <span className="w-6 h-6 rounded-full bg-[#32373C] text-white text-xs flex items-center justify-center font-bold">★</span>
        Comparison Report
      </h2>
      <p className="text-xs text-gray-500 mb-5">
        Generates a trend report from accumulated history data — showing score averages by rep and store across periods.
      </p>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
        <div>
          <label className="block text-xs font-semibold text-[#32373C] mb-1">From date</label>
          <input
            type="date"
            value={from}
            onChange={(e) => setFrom(e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C]"
          />
        </div>
        <div>
          <label className="block text-xs font-semibold text-[#32373C] mb-1">To date</label>
          <input
            type="date"
            value={to}
            onChange={(e) => setTo(e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C]"
          />
        </div>
      </div>

      {/* Period selector */}
      <div className="mb-4">
        <label className="block text-xs font-semibold text-[#32373C] mb-2">Period grouping</label>
        <div className="flex flex-wrap gap-2">
          {(['weekly', 'monthly', 'quarterly'] as CompPeriod[]).map((p) => (
            <button
              key={p}
              type="button"
              onClick={() => setPeriod(p)}
              className={`px-4 py-2 rounded-lg text-sm font-medium border transition-colors capitalize ${
                period === p
                  ? 'bg-[#32373C] text-white border-[#32373C]'
                  : 'bg-white text-[#32373C] border-gray-300 hover:border-[#32373C]'
              }`}
            >
              {p}
            </button>
          ))}
        </div>
        <p className="text-xs text-gray-400 mt-1.5">
          {period === 'weekly' && 'Columns = weeks ending on a Sunday'}
          {period === 'monthly' && 'Columns = calendar months (e.g. Jan 2026)'}
          {period === 'quarterly' && 'Columns = quarters (e.g. Q1 2026)'}
        </p>
      </div>

      {/* Filters */}
      <div className="mb-4">
        <label className="block text-xs font-semibold text-[#32373C] mb-1">
          Filters{' '}
          <span className="font-normal text-gray-400">— leave empty to include all data</span>
        </label>
        {filtersLoading && !filterOptions && (
          <p className="text-xs text-gray-400 mt-1">Loading filter options from history…</p>
        )}
        {filterOptions && (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-2">
            <MultiSelect
              label="Channel"
              items={filterOptions.channels}
              selected={selChannels}
              onChange={setSelChannels}
            />
            <MultiSelect
              label="Province"
              items={filterOptions.provinces}
              selected={selProvinces}
              onChange={setSelProvinces}
            />
            <MultiSelect
              label="Rep"
              items={filterOptions.reps.map((r) => r.email)}
              selected={selReps}
              onChange={setSelReps}
              getId={(email) => email}
              getLabel={(email) => filterOptions.reps.find((r) => r.email === email)?.name ?? email}
            />
            <MultiSelect
              label="Store"
              items={availableStoreNames}
              selected={selStores}
              onChange={setSelStores}
            />
          </div>
        )}
        {!filtersLoading && !filterOptions && (
          <p className="text-xs text-amber-600 mt-1">Could not load filter options — filters unavailable.</p>
        )}
      </div>

      {/* Delivery action */}
      <div className="mb-4">
        <label className="block text-xs font-semibold text-[#32373C] mb-2">Delivery</label>
        <div className="flex flex-wrap gap-2">
          {([
            { val: 'sharepoint', label: 'SharePoint only' },
            { val: 'email',      label: 'Email only' },
            { val: 'both',       label: 'Both' },
          ] as const).map(({ val, label }) => (
            <button
              key={val}
              type="button"
              onClick={() => { setAction(val); setEmailConfirmed(false); }}
              className={`px-4 py-2 rounded-lg text-sm font-medium border transition-colors ${
                action === val
                  ? 'bg-[#DA291C] text-white border-[#DA291C]'
                  : 'bg-white text-[#32373C] border-gray-300 hover:border-[#DA291C]'
              }`}
            >
              {label}
            </button>
          ))}
        </div>
      </div>

      {/* Email — only shown when needed */}
      {needsEmail && (
        <div className="mb-4 space-y-2">
          <div>
            <label className="block text-xs font-semibold text-[#32373C] mb-1">Email address</label>
            <input
              type="email"
              value={email}
              onChange={(e) => { setEmail(e.target.value); setEmailConfirmed(false); }}
              placeholder="recipient@example.com"
              className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C]"
            />
          </div>
          <label className="flex items-start gap-2 cursor-pointer">
            <input
              type="checkbox"
              checked={emailConfirmed}
              onChange={(e) => setEmailConfirmed(e.target.checked)}
              className="mt-0.5 accent-[#DA291C]"
            />
            <span className="text-sm text-gray-600">
              I confirm I want to send this report to <strong>{email || '…'}</strong>
            </span>
          </label>
        </div>
      )}

      <button
        type="button"
        onClick={handleRun}
        disabled={!canRun || running}
        className="px-6 py-2.5 bg-[#32373C] text-white font-bold rounded-lg text-sm hover:bg-[#1e2226] disabled:opacity-40 disabled:cursor-not-allowed transition-colors flex items-center gap-2"
      >
        {running && (
          <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
          </svg>
        )}
        {running ? 'Generating…' : 'Generate Comparison Report'}
      </button>

      {from && to && from > to && (
        <p className="mt-2 text-xs text-red-500">From date must be before to date.</p>
      )}

      {/* Result */}
      {result && (
        <div className={`mt-4 rounded-lg border p-4 text-sm ${result.error ? 'bg-red-50 border-red-200' : 'bg-green-50 border-green-200'}`}>
          {result.error ? (
            <p className="text-red-700"><strong>Error:</strong> {result.error}</p>
          ) : (
            <div className="space-y-1">
              <p className="font-semibold text-green-800">Report generated — {result.dateRange}</p>
              <p className="text-xs text-gray-600">{result.fileName}</p>
              {result.spUrl && (
                <a href={result.spUrl} target="_blank" rel="noreferrer" className="text-xs text-[#DA291C] underline block">
                  View on SharePoint
                </a>
              )}
              {result.spError && <p className="text-xs text-amber-700">SharePoint error: {result.spError}</p>}
              {result.emailSent && <p className="text-xs text-green-700">Email sent ✓</p>}
              {result.emailError && <p className="text-xs text-red-600">Email failed: {result.emailError}</p>}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ── Types ─────────────────────────────────────────────────────────────────────

interface ParsedMeta {
  reps: Rep[];
  channels: string[];
  dateRange: { from: string; to: string };
  rowCount: number;
  newEmails?: { dataEmail: string; displayName?: string }[];
}

interface GenerateResult {
  repName: string;
  fileName: string;
  status: 'ok' | 'error';
  spUrl?: string;
  spError?: string;
  emailSent?: boolean;
  repEmailSent?: boolean;
  repEmailError?: string;
  error?: string;
}

// ── MultiSelect dropdown ───────────────────────────────────────────────────────

function MultiSelect({
  label, items, selected, onChange, getId, getLabel,
}: {
  label: string;
  items: string[];
  selected: string[];
  onChange: (v: string[]) => void;
  getId?: (item: string) => string;
  getLabel?: (item: string) => string;
}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const id   = getId   ?? ((x: string) => x);
  const lbl  = getLabel ?? ((x: string) => x);
  const all  = selected.length === items.length;

  const filteredItems = query
    ? items.filter((item) => lbl(item).toLowerCase().includes(query.toLowerCase()))
    : items;

  const toggle = (item: string) => {
    onChange(
      selected.includes(id(item))
        ? selected.filter((s) => s !== id(item))
        : [...selected, id(item)]
    );
  };

  const handleOpen = () => {
    if (open) setQuery('');
    setOpen((o) => !o);
  };

  return (
    <div className="relative">
      <label className="block text-sm font-semibold text-[#32373C] mb-1">{label}</label>
      <button
        type="button"
        onClick={handleOpen}
        className="w-full flex items-center justify-between px-3 py-2 bg-white border border-gray-300 rounded-lg text-sm text-left hover:border-[#DA291C] transition-colors"
      >
        <span className="truncate">
          {selected.length === 0
            ? `Select ${label.toLowerCase()}…`
            : all
            ? `All ${label} selected`
            : `${selected.length} of ${items.length} selected`}
        </span>
        <svg className={`w-4 h-4 ml-2 shrink-0 transition-transform ${open ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
        </svg>
      </button>

      {open && (
        <div className="absolute z-20 mt-1 w-full bg-white border border-gray-200 rounded-lg shadow-lg">
          {/* Search */}
          <div className="px-3 py-2 border-b border-gray-100">
            <div className="relative">
              <svg className="absolute left-2 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-4.35-4.35M17 11A6 6 0 1 1 5 11a6 6 0 0 1 12 0z" />
              </svg>
              <input
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder={`Search ${label.toLowerCase()}…`}
                className="w-full pl-7 pr-6 py-1 text-sm border border-gray-200 rounded focus:outline-none focus:border-[#DA291C]"
                onClick={(e) => e.stopPropagation()}
              />
              {query && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  className="absolute right-1.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                >
                  <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                  </svg>
                </button>
              )}
            </div>
          </div>
          <div className="max-h-56 overflow-y-auto">
            {/* Select All */}
            <button
              type="button"
              onClick={() => onChange(all ? [] : items.map(id))}
              className="w-full flex items-center gap-2 px-3 py-2 text-sm font-semibold text-[#DA291C] hover:bg-red-50 border-b border-gray-100"
            >
              <span className={`w-4 h-4 rounded border-2 flex items-center justify-center shrink-0 ${all ? 'bg-[#DA291C] border-[#DA291C]' : 'border-gray-400'}`}>
                {all && <svg className="w-3 h-3 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" /></svg>}
              </span>
              Select All
            </button>
            {filteredItems.length === 0 ? (
              <p className="px-3 py-3 text-sm text-gray-400 text-center">No results</p>
            ) : (
              filteredItems.map((item) => {
                const checked = selected.includes(id(item));
                return (
                  <button
                    key={id(item)}
                    type="button"
                    onClick={() => toggle(item)}
                    className="w-full flex items-center gap-2 px-3 py-2 text-sm hover:bg-gray-50"
                  >
                    <span className={`w-4 h-4 rounded border-2 flex items-center justify-center shrink-0 ${checked ? 'bg-[#DA291C] border-[#DA291C]' : 'border-gray-400'}`}>
                      {checked && <svg className="w-3 h-3 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" /></svg>}
                    </span>
                    {lbl(item)}
                  </button>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Main page ─────────────────────────────────────────────────────────────────

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging]       = useState(false);
  const [file, setFile]               = useState<File | null>(null);
  const [parsing, setParsing]         = useState(false);
  const [parsedMeta, setParsedMeta]   = useState<ParsedMeta | null>(null);
  const [parseError, setParseError]   = useState<string>('');

  const [selectedReps, setSelectedReps]         = useState<string[]>([]);
  const [selectedChannels, setSelectedChannels] = useState<string[]>([]);
  const [action, setAction]                     = useState<'sharepoint' | 'email' | 'both'>('sharepoint');
  const [emailAddress, setEmailAddress]         = useState('');
  const [emailConfirmed, setEmailConfirmed]     = useState(false);
  const [emailToReps, setEmailToReps]           = useState(false);
  const [ccSelf, setCcSelf]                     = useState(false);
  const [ccEmail, setCcEmail]                   = useState('');
  const [includeExceptionReport, setIncludeExceptionReport] = useState(false);

  const [generating, setGenerating]     = useState(false);
  const [results, setResults]           = useState<GenerateResult[] | null>(null);
  const [generateError, setGenerateError] = useState('');
  const [folderName, setFolderName]     = useState('');
  const [historyAdded, setHistoryAdded] = useState<number | null>(null);
  const [historyError, setHistoryError] = useState('');

  // Exception report result
  const [exceptionReport, setExceptionReport] = useState<{
    fileName: string;
    exceptionCount: number;
    sheetSummary: { name: string; count: number }[];
    spUrl?: string;
    spError?: string;
    emailSent?: boolean;
    emailError?: string;
  } | null>(null);

  // ── File handling ──────────────────────────────────────────────────────────

  const handleFile = useCallback(async (f: File) => {
    if (!f.name.endsWith('.xlsx') && !f.name.endsWith('.xls')) {
      setParseError('Please upload an Excel file (.xlsx or .xls).');
      return;
    }
    setFile(f);
    setParsedMeta(null);
    setParseError('');
    setResults(null);
    setParsing(true);

    const fd = new FormData();
    fd.append('file', f);
    try {
      const res  = await fetch('/api/parse', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Parse failed');
      setParsedMeta(data);
      setSelectedReps(data.reps.map((r: Rep) => r.email));
      setSelectedChannels(data.channels);
    } catch (err: unknown) {
      setParseError(err instanceof Error ? err.message : 'Unknown error');
    } finally {
      setParsing(false);
    }
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const f = e.dataTransfer.files[0];
      if (f) handleFile(f);
    },
    [handleFile]
  );

  // ── Validation ─────────────────────────────────────────────────────────────

  const needsEmail   = action === 'email' || action === 'both';
  const emailValid   = !needsEmail || (emailAddress.includes('@') && emailAddress.includes('.'));
  const canSubmit    =
    !!file &&
    !!parsedMeta &&
    selectedReps.length > 0 &&
    selectedChannels.length > 0 &&
    emailValid &&
    (!needsEmail || emailConfirmed);

  // ── File size warning ──────────────────────────────────────────────────────

  const estimatedSizeMB = selectedReps.length * 0.2; // rough 200 KB per rep
  const showSizeWarning = needsEmail && selectedReps.length > 5;

  // ── Submit ─────────────────────────────────────────────────────────────────

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!canSubmit || !file) return;
    setGenerating(true);
    setResults(null);
    setGenerateError('');
    setHistoryAdded(null);
    setHistoryError('');
    setExceptionReport(null);

    const fd = new FormData();
    fd.append('file', file);
    fd.append('selectedReps',     JSON.stringify(selectedReps));
    fd.append('selectedChannels', JSON.stringify(selectedChannels));
    fd.append('action',           action);
    if (needsEmail) fd.append('emailAddress', emailAddress);
    if (emailToReps) fd.append('emailToReps', 'true');
    if (emailToReps && ccSelf && ccEmail) fd.append('ccEmail', ccEmail);
    if (includeExceptionReport) fd.append('includeExceptionReport', 'true');

    try {
      const res  = await fetch('/api/generate', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Generation failed');
      setResults(data.results);
      setFolderName(data.folderName ?? '');
      setHistoryAdded(data.historyAdded ?? null);
      setHistoryError(data.historyError ?? '');
      setExceptionReport(data.exceptionReport ?? null);
    } catch (err: unknown) {
      setGenerateError(err instanceof Error ? err.message : 'Unknown error');
    } finally {
      setGenerating(false);
    }
  };

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header className="bg-white shadow-sm border-b-4 border-[#DA291C]">
        <div className="max-w-4xl mx-auto px-6 py-4 flex items-center justify-between">
          <Image src="/vital-logo.png" alt="Vital" width={120} height={48} className="object-contain" />
          <div className="flex items-center gap-6">
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
      <main className="flex-1 max-w-4xl mx-auto w-full px-6 py-10">

        {/* Welcome banner */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-8">
          <h1 className="text-2xl font-bold text-[#DA291C] mb-2">Vital Perigee Score Card Builder</h1>
          <p className="text-sm text-gray-600 leading-relaxed">
            Welcome to the Vital Perigee Score Card Builder. Before you begin, ensure you have exported
            your form data from the{' '}
            <a href="https://live.perigeeportal.co.za/" target="_blank" rel="noreferrer" className="text-[#DA291C] underline">Perigee Portal</a>.
            Please load the raw data as is — do not make changes to the exported data.
            Make your selections and then hit <strong>Submit</strong>!
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-6">

          {/* Step 1: Upload */}
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
            <h2 className="text-base font-bold text-[#32373C] mb-4 flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-[#DA291C] text-white text-xs flex items-center justify-center font-bold">1</span>
              Upload Raw Perigee Export
            </h2>

            <div
              onClick={() => fileInputRef.current?.click()}
              onDrop={onDrop}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              className={`cursor-pointer border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
                dragging ? 'border-[#DA291C] bg-red-50' : file ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-[#DA291C]'
              }`}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }}
              />
              {parsing ? (
                <div className="flex flex-col items-center gap-2 text-[#DA291C]">
                  <svg className="animate-spin w-8 h-8" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
                  </svg>
                  <span className="text-sm font-medium">Analysing file…</span>
                </div>
              ) : file ? (
                <div className="flex flex-col items-center gap-1">
                  <svg className="w-8 h-8 text-green-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                  </svg>
                  <p className="text-sm font-semibold text-green-700">{file.name}</p>
                  {parsedMeta && (
                    <p className="text-xs text-gray-500">
                      {parsedMeta.rowCount} visits · {parsedMeta.reps.length} reps · {parsedMeta.channels.length} channels · {parsedMeta.dateRange.from} – {parsedMeta.dateRange.to}
                    </p>
                  )}
                  <p className="text-xs text-[#DA291C] underline mt-1">Click to replace</p>
                </div>
              ) : (
                <>
                  <svg className="mx-auto w-10 h-10 text-gray-400 mb-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                  <p className="text-sm font-medium text-gray-600">Drag &amp; drop your Perigee export here</p>
                  <p className="text-xs text-gray-400 mt-1">or click to browse · .xlsx only</p>
                </>
              )}
            </div>
            {parseError && <p className="mt-2 text-sm text-red-600">{parseError}</p>}
          </div>

          {/* Steps 2-5: only show after parse */}
          {parsedMeta && (
            <>
              {/* New-email warning — shown when uploaded data contains emails not yet in the alias store */}
              {parsedMeta.newEmails && parsedMeta.newEmails.length > 0 && (
                <div className="bg-amber-50 border border-amber-300 rounded-xl p-5">
                  <div className="flex items-start gap-3">
                    <div className="flex-shrink-0 w-8 h-8 rounded-full bg-amber-400 text-white flex items-center justify-center font-bold text-sm">!</div>
                    <div className="flex-1 min-w-0">
                      <h3 className="text-sm font-bold text-amber-900">
                        {parsedMeta.newEmails.length} new email{parsedMeta.newEmails.length === 1 ? '' : 's'} detected in this data
                      </h3>
                      <p className="text-xs text-amber-800 mt-1">
                        These rep email addresses don&apos;t have an alias set yet. Add an alias so each rep&apos;s report is also sent to their current address.
                      </p>
                      <div className="mt-3 max-h-28 overflow-y-auto rounded border border-amber-200 bg-white/60 divide-y divide-amber-100">
                        {parsedMeta.newEmails.map((e) => (
                          <div key={e.dataEmail} className="px-3 py-1.5 text-xs">
                            <span className="font-semibold text-amber-900">{e.displayName || '(no name)'}</span>
                            <span className="text-amber-700"> · {e.dataEmail}</span>
                          </div>
                        ))}
                      </div>
                      <Link
                        href="/admin/aliases"
                        className="inline-block mt-3 px-4 py-1.5 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-lg transition"
                      >
                        Manage Aliases →
                      </Link>
                    </div>
                  </div>
                </div>
              )}

              {/* Step 2: Reps */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
                <h2 className="text-base font-bold text-[#32373C] mb-4 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-full bg-[#DA291C] text-white text-xs flex items-center justify-center font-bold">2</span>
                  Select Reps
                </h2>
                <MultiSelect
                  label="Reps"
                  items={parsedMeta.reps.map((r) => r.email)}
                  selected={selectedReps}
                  onChange={setSelectedReps}
                  getId={(email) => email}
                  getLabel={(email) => parsedMeta.reps.find((r) => r.email === email)?.fullName ?? email}
                />
              </div>

              {/* Step 3: Channels */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
                <h2 className="text-base font-bold text-[#32373C] mb-4 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-full bg-[#DA291C] text-white text-xs flex items-center justify-center font-bold">3</span>
                  Select Channels
                </h2>
                <MultiSelect
                  label="Channels"
                  items={parsedMeta.channels}
                  selected={selectedChannels}
                  onChange={setSelectedChannels}
                />
              </div>

              {/* Step 4: Action */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
                <h2 className="text-base font-bold text-[#32373C] mb-4 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-full bg-[#DA291C] text-white text-xs flex items-center justify-center font-bold">4</span>
                  Delivery Options
                </h2>
                <div className="flex flex-wrap gap-3 mb-4">
                  {([
                    { val: 'sharepoint', label: 'SharePoint only' },
                    { val: 'email',      label: 'Email only' },
                    { val: 'both',       label: 'Both' },
                  ] as const).map(({ val, label }) => (
                    <button
                      key={val}
                      type="button"
                      onClick={() => { setAction(val); setEmailConfirmed(false); }}
                      className={`px-4 py-2 rounded-lg text-sm font-medium border transition-colors ${
                        action === val
                          ? 'bg-[#DA291C] text-white border-[#DA291C]'
                          : 'bg-white text-[#32373C] border-gray-300 hover:border-[#DA291C]'
                      }`}
                    >
                      {label}
                    </button>
                  ))}
                </div>

                {needsEmail && (
                  <div className="space-y-3">
                    <div>
                      <label className="block text-sm font-semibold mb-1">Email address</label>
                      <input
                        type="email"
                        value={emailAddress}
                        onChange={(e) => { setEmailAddress(e.target.value); setEmailConfirmed(false); }}
                        placeholder="recipient@example.com"
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C]"
                      />
                    </div>

                    {showSizeWarning && (
                      <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 text-sm text-amber-800">
                        <strong>File size notice:</strong> You&apos;ve selected {selectedReps.length} reps.
                        The total email attachment may be approximately <strong>{estimatedSizeMB.toFixed(1)} MB</strong>.
                        Large emails may be rejected by some mail servers.
                        Consider using SharePoint instead for large batches.
                      </div>
                    )}

                    <label className="flex items-start gap-2 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={emailConfirmed}
                        onChange={(e) => setEmailConfirmed(e.target.checked)}
                        className="mt-0.5 accent-[#DA291C]"
                      />
                      <span className="text-sm text-gray-600">
                        I confirm I want to send {selectedReps.length} report(s) to <strong>{emailAddress || '…'}</strong>
                      </span>
                    </label>
                  </div>
                )}

                {/* Email to Reps */}
                <div className="mt-4 pt-4 border-t border-gray-100 space-y-3">
                  <label className="flex items-start gap-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={emailToReps}
                      onChange={(e) => setEmailToReps(e.target.checked)}
                      className="mt-0.5 accent-[#DA291C]"
                    />
                    <div>
                      <span className="text-sm font-semibold text-[#32373C]">Email to Reps</span>
                      <p className="text-xs text-gray-500 mt-0.5">
                        Send each rep their own report to the email address in the data. The email is personalised with their name and score summary.
                      </p>
                      <p className="text-xs text-gray-500 mt-1">
                        If a rep has an alias set, both addresses will receive the email.{' '}
                        <Link href="/admin/aliases" className="text-[#DA291C] underline">Manage aliases</Link>.
                      </p>
                    </div>
                  </label>

                  {emailToReps && (
                    <div className="ml-6 space-y-2">
                      <label className="flex items-start gap-2 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={ccSelf}
                          onChange={(e) => setCcSelf(e.target.checked)}
                          className="mt-0.5 accent-[#DA291C]"
                        />
                        <span className="text-sm text-gray-700">Include me on each rep email (CC)</span>
                      </label>
                      {ccSelf && (
                        <input
                          type="email"
                          value={ccEmail}
                          onChange={(e) => setCcEmail(e.target.value)}
                          placeholder="your-email@example.com"
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-[#DA291C]"
                        />
                      )}
                    </div>
                  )}
                </div>

                {/* Exception Report */}
                <div className="mt-4 pt-4 border-t border-gray-100">
                  <label className="flex items-start gap-2 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={includeExceptionReport}
                      onChange={(e) => setIncludeExceptionReport(e.target.checked)}
                      className="mt-0.5 accent-[#DA291C]"
                    />
                    <div>
                      <span className="text-sm font-semibold text-[#32373C]">Generate Exception Report</span>
                      <p className="text-xs text-gray-500 mt-0.5">
                        Creates a separate Excel file listing all visits where questions were answered negatively — one sheet per question with store details, SKUs, and photo links.
                        Delivered via the same action selected above (SharePoint / Email / Both).
                      </p>
                    </div>
                  </label>
                </div>
              </div>

              {/* Step 5: Submit */}
              <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6">
                <h2 className="text-base font-bold text-[#32373C] mb-4 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-full bg-[#DA291C] text-white text-xs flex items-center justify-center font-bold">5</span>
                  Generate Reports
                </h2>

                <div className="flex items-center gap-3 flex-wrap">
                  <button
                    type="submit"
                    disabled={!canSubmit || generating}
                    className="px-8 py-3 bg-[#DA291C] text-white font-bold rounded-lg text-sm hover:bg-red-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors flex items-center gap-2"
                  >
                    {generating && (
                      <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
                      </svg>
                    )}
                    {generating ? 'Generating…' : `Generate ${selectedReps.length} Report(s)`}
                  </button>

                </div>

                {!canSubmit && parsedMeta && !generating && (
                  <p className="mt-2 text-xs text-gray-400">
                    {selectedReps.length === 0 ? 'Select at least one rep. ' : ''}
                    {selectedChannels.length === 0 ? 'Select at least one channel. ' : ''}
                    {needsEmail && !emailAddress ? 'Enter an email address. ' : ''}
                    {needsEmail && emailAddress && !emailValid ? 'Enter a valid email address. ' : ''}
                    {needsEmail && emailValid && !emailConfirmed ? 'Confirm the email send above. ' : ''}
                  </p>
                )}
              </div>
            </>
          )}
        </form>

        {/* Comparison Report (always visible) */}
        <ComparisonReport />

        {/* Generate error */}
        {generateError && (
          <div className="mt-4 bg-red-50 border border-red-200 rounded-lg p-4 text-sm text-red-700">
            <strong>Error:</strong> {generateError}
          </div>
        )}

        {/* Results */}
        {results && (
          <div className="mt-6 bg-white rounded-xl border border-gray-200 shadow-sm p-6">
            <h2 className="text-base font-bold text-[#32373C] mb-1">Results</h2>
            {folderName && (
              <p className="text-xs text-gray-500 mb-1">Folder: <code className="bg-gray-100 px-1 rounded">{folderName}</code></p>
            )}
            {historyAdded !== null && (
              <div className={`mb-4 px-3 py-2 rounded-lg border text-sm font-medium ${
                historyError
                  ? 'bg-red-50 border-red-300 text-red-700'
                  : historyAdded === 0
                  ? 'bg-amber-50 border-amber-300 text-amber-700'
                  : 'bg-green-50 border-green-300 text-green-700'
              }`}>
                {historyError
                  ? `History update failed: ${historyError}`
                  : historyAdded === 0
                  ? 'History: 0 new visits added (all already on record)'
                  : `History updated — ${historyAdded} new visit${historyAdded !== 1 ? 's' : ''} saved to SharePoint`}
              </div>
            )}
            <div className="space-y-2">
              {results.map((r, i) => (
                <div
                  key={i}
                  className={`flex items-start gap-3 p-3 rounded-lg border text-sm ${
                    r.status === 'ok' ? 'bg-green-50 border-green-200' : 'bg-red-50 border-red-200'
                  }`}
                >
                  {r.status === 'ok' ? (
                    <svg className="w-5 h-5 text-green-600 shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                    </svg>
                  ) : (
                    <svg className="w-5 h-5 text-red-600 shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                    </svg>
                  )}
                  <div>
                    <p className="font-semibold text-[#32373C]">{r.repName}</p>
                    <p className="text-xs text-gray-500">{r.fileName}</p>
                    {r.spUrl && (
                      <a href={r.spUrl} target="_blank" rel="noreferrer" className="text-xs text-[#DA291C] underline">
                        View on SharePoint
                      </a>
                    )}
                    {r.spError && <p className="text-xs text-amber-600 mt-0.5">SharePoint: {r.spError}</p>}
                    {r.emailSent && <p className="text-xs text-green-600 mt-0.5">Summary email sent ✓</p>}
                    {r.repEmailSent && <p className="text-xs text-green-600 mt-0.5">Rep email sent ✓</p>}
                    {r.repEmailError && <p className="text-xs text-amber-600 mt-0.5">Rep email: {r.repEmailError}</p>}
                    {r.error && <p className="text-xs text-red-600 mt-0.5">{r.error}</p>}
                  </div>
                </div>
              ))}
            </div>

            {/* Exception Report result */}
            {exceptionReport && (
              <div className="mt-4 pt-4 border-t border-gray-200">
                <h3 className="text-sm font-bold text-[#32373C] mb-2">Exception Report</h3>
                <div className={`p-3 rounded-lg border text-sm ${
                  exceptionReport.spError && !exceptionReport.spUrl ? 'bg-red-50 border-red-200' : 'bg-green-50 border-green-200'
                }`}>
                  <p className="font-semibold text-[#32373C]">
                    {exceptionReport.exceptionCount} exception{exceptionReport.exceptionCount !== 1 ? 's' : ''} across {exceptionReport.sheetSummary.length} question{exceptionReport.sheetSummary.length !== 1 ? 's' : ''}
                  </p>
                  <p className="text-xs text-gray-500">{exceptionReport.fileName}</p>
                  {exceptionReport.sheetSummary.length > 0 && (
                    <ul className="mt-1 space-y-0.5">
                      {exceptionReport.sheetSummary.map((s) => (
                        <li key={s.name} className="text-xs text-gray-600">
                          <span className="font-medium text-[#DA291C]">{s.count}</span> — {s.name}
                        </li>
                      ))}
                    </ul>
                  )}
                  {exceptionReport.spUrl && (
                    <a href={exceptionReport.spUrl} target="_blank" rel="noreferrer" className="text-xs text-[#DA291C] underline block mt-1">
                      View on SharePoint
                    </a>
                  )}
                  {exceptionReport.spError && <p className="text-xs text-amber-600 mt-0.5">SharePoint: {exceptionReport.spError}</p>}
                  {exceptionReport.emailSent && <p className="text-xs text-green-600 mt-0.5">Exception report email sent ✓</p>}
                  {exceptionReport.emailError && <p className="text-xs text-red-600 mt-0.5">Email failed: {exceptionReport.emailError}</p>}
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 py-4 text-center text-xs text-gray-400">
        If you have any suggestions about how we might improve this tool, please email{' '}
        <a href="mailto:info@outerjoin.co.za" className="text-[#E8572A] underline">
          info@outerjoin.co.za
        </a>
      </footer>
    </div>
  );
}
