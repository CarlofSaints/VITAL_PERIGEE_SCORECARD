'use client';

import { useState, useRef, useCallback } from 'react';
import Image from 'next/image';
import type { Rep } from '@/types';

// ── Types ─────────────────────────────────────────────────────────────────────

interface ParsedMeta {
  reps: Rep[];
  channels: string[];
  dateRange: { from: string; to: string };
  rowCount: number;
}

interface GenerateResult {
  repName: string;
  fileName: string;
  status: 'ok' | 'error';
  spUrl?: string;
  emailSent?: boolean;
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
  const id   = getId   ?? ((x: string) => x);
  const lbl  = getLabel ?? ((x: string) => x);
  const all  = selected.length === items.length;

  const toggle = (item: string) => {
    onChange(
      selected.includes(id(item))
        ? selected.filter((s) => s !== id(item))
        : [...selected, id(item)]
    );
  };

  return (
    <div className="relative">
      <label className="block text-sm font-semibold text-[#32373C] mb-1">{label}</label>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
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
        <div className="absolute z-20 mt-1 w-full bg-white border border-gray-200 rounded-lg shadow-lg max-h-64 overflow-y-auto">
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
          {items.map((item) => {
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
          })}
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

  const [generating, setGenerating]     = useState(false);
  const [results, setResults]           = useState<GenerateResult[] | null>(null);
  const [generateError, setGenerateError] = useState('');
  const [folderName, setFolderName]     = useState('');

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

    const fd = new FormData();
    fd.append('file', file);
    fd.append('selectedReps',     JSON.stringify(selectedReps));
    fd.append('selectedChannels', JSON.stringify(selectedChannels));
    fd.append('action',           action);
    if (needsEmail) fd.append('emailAddress', emailAddress);

    try {
      const res  = await fetch('/api/generate', { method: 'POST', body: fd });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error ?? 'Generation failed');
      setResults(data.results);
      setFolderName(data.folderName ?? '');
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
          <Image src="/outerjoin-logo.png" alt="OuterJoin" width={160} height={48} className="object-contain" />
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 max-w-4xl mx-auto w-full px-6 py-10">

        {/* Welcome banner */}
        <div className="bg-white rounded-xl border border-gray-200 shadow-sm p-6 mb-8">
          <h1 className="text-2xl font-bold text-[#DA291C] mb-2">Vital Perigee Score Card Builder</h1>
          <p className="text-sm text-gray-600 leading-relaxed">
            Welcome to the Vital Perigee Score Card Builder. Before you begin, ensure you have exported
            your form data from the Perigee portal.
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

                  {/* Placeholder: Summary Report */}
                  <button
                    type="button"
                    disabled
                    title="Coming soon"
                    className="px-6 py-3 bg-gray-100 text-gray-400 font-medium rounded-lg text-sm border border-gray-200 cursor-not-allowed"
                  >
                    Run Summary Report <span className="text-xs">(coming soon)</span>
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
              <p className="text-xs text-gray-500 mb-4">Folder: <code className="bg-gray-100 px-1 rounded">{folderName}</code></p>
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
                    {r.emailSent && <span className="text-xs text-green-600 ml-2">Email sent ✓</span>}
                    {r.error && <p className="text-xs text-red-600 mt-0.5">{r.error}</p>}
                  </div>
                </div>
              ))}
            </div>
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
