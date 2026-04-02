'use client';

import { useState, useEffect } from 'react';
import Image from 'next/image';
import Link from 'next/link';
import type { LogEntry } from '@/lib/activityLog';

function formatTimestamp(iso: string): string {
  const d = new Date(iso);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const date = `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;
  const time = `${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}`;
  return `${date} · ${time}`;
}

function Badge({ label, color }: { label: string; color: 'red' | 'gray' | 'green' | 'amber' }) {
  const cls = {
    red:   'bg-red-100 text-red-700 border-red-200',
    gray:  'bg-gray-100 text-gray-600 border-gray-200',
    green: 'bg-green-100 text-green-700 border-green-200',
    amber: 'bg-amber-100 text-amber-700 border-amber-200',
  }[color];
  return (
    <span className={`inline-block text-xs font-semibold px-2 py-0.5 rounded border ${cls}`}>
      {label}
    </span>
  );
}

function SuccessBadge({ ok }: { ok: boolean }) {
  return ok
    ? <span className="inline-flex items-center gap-1 text-xs font-semibold text-green-700"><span className="w-4 h-4 rounded-full bg-green-100 flex items-center justify-center">✓</span>Success</span>
    : <span className="inline-flex items-center gap-1 text-xs font-semibold text-red-600"><span className="w-4 h-4 rounded-full bg-red-100 flex items-center justify-center">✗</span>Failed</span>;
}

function EntryRow({ entry }: { entry: LogEntry }) {
  const [open, setOpen] = useState(false);
  const isGenerate = entry.type === 'generate';

  const spCount = entry.reports?.filter(r => r.spSuccess).length ?? 0;
  const spTotal = entry.reports?.length ?? 0;
  const repEmailCount = entry.reports?.filter(r => r.repEmailSent).length ?? 0;

  return (
    <>
      <tr
        className="border-b border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors"
        onClick={() => setOpen(o => !o)}
      >
        {/* Type */}
        <td className="px-4 py-3 whitespace-nowrap">
          <Badge
            label={isGenerate ? 'Generate' : 'Comparison'}
            color={isGenerate ? 'red' : 'gray'}
          />
        </td>

        {/* Date range */}
        <td className="px-4 py-3 text-sm text-[#32373C] max-w-[200px]">
          <span className="font-medium">{entry.dateRange}</span>
          {entry.period && <span className="block text-xs text-gray-400 mt-0.5">{entry.period}</span>}
        </td>

        {/* Action */}
        <td className="px-4 py-3 whitespace-nowrap">
          <Badge
            label={entry.action === 'both' ? 'SP + Email' : entry.action === 'email' ? 'Email only' : 'SP only'}
            color={entry.action === 'both' ? 'red' : entry.action === 'email' ? 'amber' : 'gray'}
          />
        </td>

        {/* SharePoint */}
        <td className="px-4 py-3 text-sm text-center">
          {isGenerate ? (
            <span className={`font-medium ${spCount === spTotal ? 'text-green-700' : 'text-amber-600'}`}>
              {spCount}/{spTotal}
            </span>
          ) : (
            entry.compSpUrl
              ? <span className="text-green-700 font-medium">✓</span>
              : entry.compSpError
              ? <span className="text-red-600 font-medium">✗</span>
              : <span className="text-gray-400">—</span>
          )}
        </td>

        {/* Summary email */}
        <td className="px-4 py-3 text-sm text-center">
          {isGenerate ? (
            entry.summaryEmail
              ? entry.summaryEmail.sent
                ? <span className="text-green-700 text-xs font-medium">✓ {entry.summaryEmail.to}</span>
                : <span className="text-red-600 text-xs font-medium">✗ failed</span>
              : <span className="text-gray-400">—</span>
          ) : (
            entry.compEmailSent
              ? <span className="text-green-700 text-xs font-medium">✓ {entry.compEmailTo}</span>
              : entry.compEmailError
              ? <span className="text-red-600 text-xs font-medium">✗ failed</span>
              : <span className="text-gray-400">—</span>
          )}
        </td>

        {/* Rep emails */}
        <td className="px-4 py-3 text-sm text-center">
          {isGenerate && entry.reports?.some(r => r.repEmailSent !== undefined) ? (
            <span className={`font-medium ${repEmailCount === spTotal ? 'text-green-700' : 'text-amber-600'}`}>
              {repEmailCount}/{spTotal}
            </span>
          ) : (
            <span className="text-gray-400">—</span>
          )}
        </td>

        {/* History */}
        <td className="px-4 py-3 text-sm text-center">
          {entry.historyAdded !== undefined ? (
            <span className={entry.historyAdded > 0 ? 'text-green-700 font-medium' : 'text-gray-400'}>
              +{entry.historyAdded}
            </span>
          ) : <span className="text-gray-400">—</span>}
        </td>

        {/* Status */}
        <td className="px-4 py-3 whitespace-nowrap">
          <SuccessBadge ok={entry.overallSuccess} />
        </td>

        {/* Time */}
        <td className="px-4 py-3 text-xs text-gray-400 whitespace-nowrap">
          {formatTimestamp(entry.timestamp)}
        </td>

        {/* Expand toggle */}
        <td className="px-4 py-3 text-center">
          <svg
            className={`w-4 h-4 text-gray-400 transition-transform inline-block ${open ? 'rotate-180' : ''}`}
            fill="none" viewBox="0 0 24 24" stroke="currentColor"
          >
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
          </svg>
        </td>
      </tr>

      {open && (
        <tr className="bg-gray-50 border-b border-gray-100">
          <td colSpan={10} className="px-6 py-4">
            {isGenerate && entry.reports && entry.reports.length > 0 ? (
              <div className="space-y-2">
                <p className="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-3">Per-Rep Detail</p>
                {entry.reports.map((r, i) => (
                  <div key={i} className={`flex flex-wrap items-start gap-3 p-3 rounded-lg border text-sm ${r.success ? 'bg-white border-gray-200' : 'bg-red-50 border-red-200'}`}>
                    <div className="min-w-[160px]">
                      <p className="font-semibold text-[#32373C]">{r.repName}</p>
                      <p className="text-xs text-gray-400 break-all">{r.fileName}</p>
                    </div>
                    <div className="flex flex-wrap gap-2 items-center text-xs">
                      {r.spSuccess
                        ? <span className="text-green-700">✓ SharePoint{r.spUrl ? <> · <a href={r.spUrl} target="_blank" rel="noreferrer" className="underline text-[#DA291C]">View</a></> : ''}</span>
                        : r.spError
                        ? <span className="text-amber-600">⚠ SP: {r.spError}</span>
                        : null
                      }
                      {r.repEmailSent !== undefined && (
                        r.repEmailSent
                          ? <span className="text-green-700">✓ Email → {r.repEmailTo}{r.repEmailCc ? ` (CC: ${r.repEmailCc})` : ''}</span>
                          : <span className="text-red-600">✗ Email failed: {r.repEmailError}</span>
                      )}
                      {r.error && <span className="text-red-600">{r.error}</span>}
                    </div>
                  </div>
                ))}
                {entry.historyError && (
                  <p className="text-xs text-amber-700 mt-2">History error: {entry.historyError}</p>
                )}
              </div>
            ) : (
              <div className="text-sm space-y-1">
                {entry.compFileName && <p className="text-gray-600"><span className="font-medium">File:</span> {entry.compFileName}</p>}
                {entry.compSpUrl && <p><a href={entry.compSpUrl} target="_blank" rel="noreferrer" className="text-[#DA291C] underline text-xs">View on SharePoint</a></p>}
                {entry.compSpError && <p className="text-amber-700 text-xs">SP error: {entry.compSpError}</p>}
                {entry.compEmailError && <p className="text-red-600 text-xs">Email error: {entry.compEmailError}</p>}
              </div>
            )}
          </td>
        </tr>
      )}
    </>
  );
}

export default function LogPage() {
  const [entries, setEntries] = useState<LogEntry[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState('');

  useEffect(() => {
    fetch('/api/log')
      .then(r => r.json())
      .then(data => {
        if (data.error) setError(data.error);
        else setEntries(data.entries ?? []);
      })
      .catch(err => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header className="bg-white shadow-sm border-b-4 border-[#DA291C]">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <Image src="/vital-logo.png" alt="Vital" width={120} height={48} className="object-contain" />
          <div className="flex items-center gap-6">
            <Link href="/" className="text-sm text-[#32373C] hover:text-[#DA291C] font-medium transition-colors">
              ← Score Card Builder
            </Link>
            <Image src="/perigee-logo.jpg" alt="Perigee" width={80} height={48} className="object-contain" />
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 max-w-6xl mx-auto w-full px-6 py-10">
        <div className="mb-6">
          <h1 className="text-2xl font-bold text-[#DA291C] mb-1">Activity Log</h1>
          <p className="text-sm text-gray-500">Record of all score card generations, comparisons, and emails. Most recent first.</p>
        </div>

        {loading && (
          <div className="flex items-center gap-3 text-gray-500 py-10 justify-center">
            <svg className="animate-spin w-6 h-6 text-[#DA291C]" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
            </svg>
            Loading activity log…
          </div>
        )}

        {error && (
          <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-sm text-red-700">
            <strong>Error loading log:</strong> {error}
          </div>
        )}

        {!loading && !error && entries.length === 0 && (
          <div className="text-center py-16 text-gray-400">
            <p className="text-lg font-medium mb-1">No activity yet</p>
            <p className="text-sm">Entries will appear here after you generate your first score card batch.</p>
          </div>
        )}

        {!loading && !error && entries.length > 0 && (
          <div className="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead>
                  <tr className="border-b border-gray-200 bg-[#32373C]">
                    {['Type', 'Date Range', 'Action', 'SharePoint', 'Summary Email', 'Rep Emails', 'History', 'Status', 'Time', ''].map((h) => (
                      <th key={h} className="px-4 py-3 text-xs font-semibold text-white uppercase tracking-wide whitespace-nowrap">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {entries.map((entry) => (
                    <EntryRow key={entry.id} entry={entry} />
                  ))}
                </tbody>
              </table>
            </div>
            <div className="px-4 py-2 border-t border-gray-100 bg-gray-50 text-xs text-gray-400 text-right">
              {entries.length} entr{entries.length !== 1 ? 'ies' : 'y'} · Click any row to expand
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 py-4 text-center text-xs text-gray-400">
        If you have any suggestions about how we might improve this tool, please email{' '}
        <a href="mailto:info@outerjoin.co.za" className="text-[#E8572A] underline">info@outerjoin.co.za</a>
      </footer>
    </div>
  );
}
