/**
 * Vital Score Card — Activity Log
 *
 * Stored as JSON on SharePoint:
 *   Historical Data/VITAL_ACTIVITY_LOG.json
 *
 * Max 200 entries — oldest are trimmed automatically.
 */

import { downloadLogFile, uploadLogFile } from '@/lib/sharepoint';

export interface LogReport {
  repName: string;
  fileName: string;
  spSuccess: boolean;
  spUrl?: string;
  spError?: string;
  repEmailSent?: boolean;
  repEmailTo?: string;
  repEmailCc?: string;
  repEmailError?: string;
  success: boolean;
  error?: string;
}

export interface LogEntry {
  id: string;
  timestamp: string;          // ISO string
  type: 'generate' | 'comparison';
  dateRange: string;          // e.g. "01 Mar 2026 – 31 Mar 2026"
  action: string;             // 'sharepoint' | 'email' | 'both'

  // Generate-specific
  reports?: LogReport[];
  summaryEmail?: {
    sent: boolean;
    to: string;
    attachmentCount: number;
    error?: string;
  };
  historyAdded?: number;
  historyError?: string;

  // Comparison-specific
  period?: string;
  compFileName?: string;
  compSpUrl?: string;
  compSpError?: string;
  compEmailSent?: boolean;
  compEmailTo?: string;
  compEmailError?: string;

  overallSuccess: boolean;
}

const MAX_ENTRIES = 200;

function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

export async function readActivityLog(): Promise<LogEntry[]> {
  try {
    const buf = await downloadLogFile();
    if (!buf) return [];
    return JSON.parse(buf.toString('utf-8')) as LogEntry[];
  } catch {
    return [];
  }
}

export async function appendActivityLog(
  entry: Omit<LogEntry, 'id' | 'timestamp'>
): Promise<void> {
  const entries = await readActivityLog();
  const newEntry: LogEntry = {
    id: generateId(),
    timestamp: new Date().toISOString(),
    ...entry,
  };
  entries.unshift(newEntry); // newest first
  const trimmed = entries.slice(0, MAX_ENTRIES);
  const json = JSON.stringify(trimmed, null, 2);
  await uploadLogFile(Buffer.from(json, 'utf-8'));
}
